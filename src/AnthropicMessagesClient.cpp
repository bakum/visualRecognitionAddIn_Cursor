#include "AnthropicMessagesClient.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#if defined(_WIN32)
#include <Windows.h>
#include <winhttp.h>
#endif

#include <boost/json.hpp>
#if !defined(_WIN32)
#include <boost/asio/connect.hpp>
#include <boost/asio/ip/tcp.hpp>
#include <boost/asio/ssl/context.hpp>
#include <boost/asio/ssl/stream.hpp>
#include <boost/beast/core.hpp>
#include <boost/beast/http.hpp>
#include <boost/beast/ssl.hpp>
#include <boost/beast/version.hpp>
#endif

#include <algorithm>
#include <chrono>
#include <cctype>
#include <cmath>
#include <codecvt>
#include <limits>
#include <locale>
#include <optional>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

namespace {

constexpr char kAnthropicHostUtf8[] = "api.anthropic.com";
#if defined(_WIN32)
constexpr wchar_t kAnthropicHost[] = L"api.anthropic.com";
constexpr wchar_t kAnthropicPath[] = L"/v1/messages";
#else
constexpr char kAnthropicPath[] = "/v1/messages";
#endif

constexpr char kAnthropicApiVersion[] = "2023-06-01";

constexpr int kWinHttpResolveTimeoutMs = 10000;
constexpr int kWinHttpConnectTimeoutMs = 10000;
constexpr int kWinHttpSendTimeoutMs = 30000;
constexpr int kWinHttpReceiveTimeoutMs = 45000;
constexpr int kWinHttpTotalRequestDeadlineMs = 65000;
constexpr int kWinHttpMinReceiveTimeoutMs = 1000;
constexpr int kWinHttpMinTotalDeadlineMs = 5000;

constexpr size_t kMaxPlainTextPromptBytes = 2U * 1024U * 1024U;

/// Совпадает с подсказкой Gemini для единообразия поведения в 1С.
constexpr char kPlainTextSystemInstruction[] =
    "When the user message has line breaks, indentation, lists, tables, or multiple paragraphs, treat "
    "that structure as meaningful. In your answer, preserve the same kind of formatting (line breaks, "
    "spacing, bullet/numbered lists, aligned blocks) whenever you reproduce, quote, transform, or "
    "build on their text. Do not collapse whitespace or flatten layout unless the user explicitly "
    "asks for plain or compact output. If the task involves reading text from an image or document "
    "snapshot, first infer orientation and read in corrected orientation (including 90 degrees, "
    "180 degrees, and 270 degrees rotations) when needed.";

void TrimAsciiInPlace(std::string& s) {
    while (!s.empty() && (static_cast<unsigned char>(s.front()) <= ' ')) {
        s.erase(s.begin());
    }
    while (!s.empty() && (static_cast<unsigned char>(s.back()) <= ' ')) {
        s.pop_back();
    }
}

std::string Base64EncodeBytes(const std::vector<char>& in) {
    static const char tbl[] =
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    std::string out;
    const auto* bytes = reinterpret_cast<const unsigned char*>(in.data());
    const size_t n = in.size();
    out.reserve(((n + 2U) / 3U) * 4U);
    for (size_t i = 0; i < n; i += 3U) {
        const unsigned b0 = bytes[i];
        const unsigned b1 = i + 1U < n ? bytes[i + 1U] : 0U;
        const unsigned b2 = i + 2U < n ? bytes[i + 2U] : 0U;
        const unsigned triple = (b0 << 16U) | (b1 << 8U) | b2;
        out += tbl[(triple >> 18U) & 63U];
        out += tbl[(triple >> 12U) & 63U];
        out += (i + 1U < n) ? tbl[(triple >> 6U) & 63U] : '=';
        out += (i + 2U < n) ? tbl[triple & 63U] : '=';
    }
    return out;
}

bool BytesLookLikePdf(const std::vector<char>& b) {
    return b.size() >= 5 && b[0] == '%' && b[1] == 'P' && b[2] == 'D' && b[3] == 'F';
}

bool IsPrimaryInlineMime(std::string_view mime) {
    return mime == "application/pdf" || mime == "image/jpeg" || mime == "image/png"
        || mime == "image/gif" || mime == "image/webp" || mime == "image/bmp" || mime == "image/tiff";
}

std::optional<std::string> ExtractAnthropicStopReason(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& o = v.as_object();
        const auto it = o.find("stop_reason");
        if (it == o.end() || !it->value().is_string()) {
            return std::nullopt;
        }
        return std::string(it->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}

std::wstring Utf8ToWide(std::string_view u8) {
    if (u8.empty()) {
        return {};
    }
#if defined(_WIN32)
    const int n = MultiByteToWideChar(CP_UTF8, 0, u8.data(), static_cast<int>(u8.size()), nullptr, 0);
    if (n <= 0) {
        return {};
    }
    std::wstring w(static_cast<size_t>(n), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, u8.data(), static_cast<int>(u8.size()), w.data(), n);
    return w;
#else
    static std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> cvt_utf8_utf16;
    return cvt_utf8_utf16.from_bytes(u8.data(), u8.data() + u8.size());
#endif
}

bool IsValidAnthropicModelId(std::string_view s) {
    if (s.empty() || s.size() > 160) {
        return false;
    }
    for (unsigned char uc : s) {
        const char c = static_cast<char>(uc);
        if (std::isalnum(uc) != 0 || c == '-' || c == '.' || c == '_') {
            continue;
        }
        return false;
    }
    return true;
}

std::string BuildAnthropicModelAliasKey(std::string_view value) {
    std::string key;
    key.reserve(value.size());
    for (unsigned char uc : value) {
        if (std::isalnum(uc) == 0) {
            continue;
        }
        key.push_back(static_cast<char>(std::tolower(uc)));
    }
    return key;
}

std::string ResolveAnthropicModelAlias(std::string_view model_id_utf8) {
    std::string model(model_id_utf8);
    TrimAsciiInPlace(model);
    if (model.empty()) {
        return std::string(kAnthropicDefaultModelId);
    }
    const std::string key = BuildAnthropicModelAliasKey(model);
    if (key == "claudesonnet45" || key == "sonnet45" || key == "claude45sonnet") {
        return "claude-sonnet-4-5-20250929";
    }
    if (key == "claudehaiku45" || key == "haiku45") {
        return "claude-haiku-4-5-20251001";
    }
    if (key == "claude35sonnet" || key == "sonnet35") {
        return "claude-3-5-sonnet-20241022";
    }
    if (key == "claude35haiku" || key == "haiku35") {
        return "claude-3-5-haiku-20241022";
    }
    if (key == "claude3opus" || key == "opus3") {
        return "claude-3-opus-20240229";
    }
    return model;
}

std::string BuildMessagesRequestBody(const std::string& model_resolved,
                                       const int max_output_tokens,
                                       const std::string& user_text_utf8) {
    boost::json::object user_msg;
    user_msg["role"] = "user";
    user_msg["content"] = user_text_utf8;

    boost::json::array messages;
    messages.push_back(user_msg);

    boost::json::object root;
    root["model"] = model_resolved;
    root["max_tokens"] = max_output_tokens;
    root["system"] = std::string(kPlainTextSystemInstruction);
    root["messages"] = messages;

    return std::string(boost::json::serialize(root));
}

bool HttpPostAnthropic(const std::string& api_key_utf8,
                       const std::string& body_utf8,
                       const GeminiHttpTimeouts* timeouts,
                       std::string_view anthropic_beta_value_opt,
                       unsigned long& status_out,
                       std::string& response_body_out,
                       std::string& error_out) {
#if defined(_WIN32)
    const auto is_transient_winhttp_error = [](DWORD e) -> bool {
        return e == ERROR_WINHTTP_TIMEOUT || e == ERROR_WINHTTP_NAME_NOT_RESOLVED
            || e == ERROR_WINHTTP_CANNOT_CONNECT || e == ERROR_WINHTTP_CONNECTION_ERROR
            || e == ERROR_WINHTTP_RESEND_REQUEST || e == ERROR_WINHTTP_INTERNAL_ERROR;
    };

    const int configured_receive_timeout_ms = std::max(
        kWinHttpMinReceiveTimeoutMs,
        (timeouts && timeouts->receive_timeout_ms > 0) ? timeouts->receive_timeout_ms
                                                       : kWinHttpReceiveTimeoutMs);
    const int configured_total_deadline_ms = std::max(
        kWinHttpMinTotalDeadlineMs,
        (timeouts && timeouts->total_deadline_ms > 0) ? timeouts->total_deadline_ms
                                                       : kWinHttpTotalRequestDeadlineMs);

    HINTERNET session =
        WinHttpOpen(L"VisualRecognitionAddIn/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) {
        error_out = "Anthropic API network error (WinHttpOpen)";
        return false;
    }
    WinHttpSetTimeouts(session, kWinHttpResolveTimeoutMs, kWinHttpConnectTimeoutMs,
                       kWinHttpSendTimeoutMs, configured_receive_timeout_ms);

    HINTERNET connect = WinHttpConnect(session, kAnthropicHost, INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!connect) {
        error_out = "Anthropic API network error (WinHttpConnect)";
        WinHttpCloseHandle(session);
        return false;
    }

    const std::wstring wkey = Utf8ToWide(api_key_utf8);
    const std::wstring wver = Utf8ToWide(kAnthropicApiVersion);
    std::wstring headers = L"x-api-key: " + wkey + L"\r\nanthropic-version: " + wver
        + L"\r\nContent-Type: application/json\r\n";
    if (!anthropic_beta_value_opt.empty()) {
        headers += L"anthropic-beta: " + Utf8ToWide(std::string(anthropic_beta_value_opt)) + L"\r\n";
    }

    constexpr int kMaxAttempts = 3;
    const auto started_at = std::chrono::steady_clock::now();
    bool ok = false;
    for (int attempt = 1; attempt <= kMaxAttempts; ++attempt) {
        const auto now = std::chrono::steady_clock::now();
        const auto elapsed_ms =
            std::chrono::duration_cast<std::chrono::milliseconds>(now - started_at).count();
        const int remaining_total_ms =
            configured_total_deadline_ms - static_cast<int>(elapsed_ms);
        if (remaining_total_ms <= 0) {
            error_out = "Anthropic API timeout: total request deadline exceeded";
            break;
        }
        const int receive_timeout_ms = std::min(configured_receive_timeout_ms, remaining_total_ms);

        HINTERNET request =
            WinHttpOpenRequest(connect, L"POST", kAnthropicPath, nullptr, WINHTTP_NO_REFERER,
                               WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
        if (!request) {
            const DWORD e = GetLastError();
            error_out = "Anthropic API network error (WinHttpOpenRequest, code=" + std::to_string(e)
                + ")";
            continue;
        }
        WinHttpSetTimeouts(request, kWinHttpResolveTimeoutMs, kWinHttpConnectTimeoutMs,
                           kWinHttpSendTimeoutMs, receive_timeout_ms);

        const BOOL sent =
            WinHttpSendRequest(request, headers.c_str(), static_cast<DWORD>(-1),
                               const_cast<char*>(body_utf8.data()), static_cast<DWORD>(body_utf8.size()),
                               static_cast<DWORD>(body_utf8.size()), 0);
        if (!sent) {
            const DWORD e = GetLastError();
            error_out = "Anthropic API network error (WinHttpSendRequest, code=" + std::to_string(e)
                + ")";
            WinHttpCloseHandle(request);
            if (is_transient_winhttp_error(e) && attempt < kMaxAttempts) {
                Sleep(250);
                continue;
            }
            WinHttpCloseHandle(connect);
            WinHttpCloseHandle(session);
            return false;
        }

        if (!WinHttpReceiveResponse(request, nullptr)) {
            const DWORD e = GetLastError();
            if (e == ERROR_WINHTTP_TIMEOUT) {
                error_out = "Anthropic API timeout in WinHttpReceiveResponse (code=12002)";
            } else {
                error_out = "Anthropic API network error (WinHttpReceiveResponse, code="
                    + std::to_string(e) + ")";
            }
            WinHttpCloseHandle(request);
            if (is_transient_winhttp_error(e) && attempt < kMaxAttempts) {
                Sleep(300);
                continue;
            }
            WinHttpCloseHandle(connect);
            WinHttpCloseHandle(session);
            return false;
        }

        DWORD status = 0;
        DWORD sz = sizeof(status);
        if (WinHttpQueryHeaders(request,
                                WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
                                WINHTTP_HEADER_NAME_BY_INDEX, &status, &sz, WINHTTP_NO_HEADER_INDEX)) {
            status_out = status;
        } else {
            status_out = 0;
        }

        response_body_out.clear();
        bool read_ok = true;
        for (;;) {
            DWORD avail = 0;
            if (!WinHttpQueryDataAvailable(request, &avail)) {
                const DWORD e = GetLastError();
                error_out = "Anthropic API network error (WinHttpQueryDataAvailable, code="
                    + std::to_string(e) + ")";
                read_ok = false;
                break;
            }
            if (avail == 0) {
                break;
            }
            std::vector<char> chunk(static_cast<size_t>(avail));
            DWORD read = 0;
            if (!WinHttpReadData(request, chunk.data(), avail, &read)) {
                const DWORD e = GetLastError();
                error_out = "Anthropic API network error (WinHttpReadData, code=" + std::to_string(e)
                    + ")";
                read_ok = false;
                break;
            }
            response_body_out.append(chunk.data(), read);
        }
        WinHttpCloseHandle(request);
        if (read_ok) {
            ok = true;
            break;
        }
    }

    if (!ok) {
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        return false;
    }

    WinHttpCloseHandle(connect);
    WinHttpCloseHandle(session);
    return true;
#else
    namespace beast = boost::beast;
    namespace http = beast::http;
    namespace net = boost::asio;
    namespace ssl = boost::asio::ssl;
    using tcp = net::ip::tcp;

    try {
        net::io_context ioc;
        ssl::context ctx(ssl::context::tls_client);
        ctx.set_default_verify_paths();

        tcp::resolver resolver(ioc);
        beast::ssl_stream<beast::tcp_stream> stream(ioc, ctx);
        if (!SSL_set_tlsext_host_name(stream.native_handle(), kAnthropicHostUtf8)) {
            error_out = "Anthropic API TLS error (SNI)";
            return false;
        }

        const auto results = resolver.resolve(kAnthropicHostUtf8, "443");
        beast::get_lowest_layer(stream).connect(results);
        stream.handshake(ssl::stream_base::client);

        http::request<http::string_body> req{http::verb::post, kAnthropicPath, 11};
        req.set(http::field::host, kAnthropicHostUtf8);
        req.set(http::field::user_agent, "VisualRecognitionAddIn/1.0");
        req.set("x-api-key", api_key_utf8);
        req.set("anthropic-version", kAnthropicApiVersion);
        req.set(http::field::content_type, "application/json");
        if (!anthropic_beta_value_opt.empty()) {
            req.insert("anthropic-beta", std::string(anthropic_beta_value_opt));
        }
        req.body() = body_utf8;
        req.prepare_payload();

        http::write(stream, req);

        beast::flat_buffer buffer;
        http::response<http::string_body> res;
        http::read(stream, buffer, res);
        status_out = static_cast<unsigned long>(res.result_int());
        response_body_out = std::move(res.body());

        beast::error_code ec;
        stream.shutdown(ec);
        if (ec == net::error::eof) {
            ec = {};
        }
        if (ec) {
            error_out = "Anthropic API TLS shutdown error";
            return false;
        }
        return true;
    } catch (const std::exception& e) {
        error_out = std::string("Anthropic API network error: ") + e.what();
        return false;
    }
#endif
}

std::optional<std::string> ExtractAnthropicApiErrorMessage(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto type_it = root.find("type");
        if (type_it != root.end() && type_it->value().is_string()
            && std::string(type_it->value().as_string()) == "error") {
            const auto err_it = root.find("error");
            if (err_it != root.end() && err_it->value().is_object()) {
                const boost::json::object& err = err_it->value().as_object();
                const auto msg_it = err.find("message");
                if (msg_it != err.end() && msg_it->value().is_string()) {
                    return std::string(msg_it->value().as_string());
                }
            }
        }
        const auto msg_top = root.find("message");
        if (msg_top != root.end() && msg_top->value().is_string()) {
            return std::string(msg_top->value().as_string());
        }
        return std::string("Anthropic API error");
    } catch (...) {
        return std::nullopt;
    }
}

bool IsTransientAnthropicOverload(unsigned long http_status,
                                  const std::optional<std::string>& api_err) {
    if (http_status == 429UL || http_status == 503UL || http_status == 529UL) {
        return true;
    }
    if (!api_err.has_value()) {
        return false;
    }
    std::string s = *api_err;
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    return s.find("overload") != std::string::npos || s.find("rate limit") != std::string::npos
        || s.find("try again") != std::string::npos || s.find("temporar") != std::string::npos;
}

void ExtractAnthropicUsage(const std::string& json, GeminiUsageStats* usage_out) {
    if (!usage_out) {
        return;
    }
    *usage_out = GeminiUsageStats{};
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return;
        }
        const boost::json::object& root = v.as_object();
        const auto uit = root.find("usage");
        if (uit == root.end() || !uit->value().is_object()) {
            return;
        }
        const boost::json::object& usage = uit->value().as_object();
        const auto read_i64 = [&usage](const char* key) -> std::optional<int64_t> {
            const auto it = usage.find(key);
            if (it == usage.end()) {
                return std::nullopt;
            }
            if (it->value().is_int64()) {
                return it->value().as_int64();
            }
            if (it->value().is_uint64()) {
                const uint64_t val = it->value().as_uint64();
                if (val <= static_cast<uint64_t>(std::numeric_limits<int64_t>::max())) {
                    return static_cast<int64_t>(val);
                }
            }
            return std::nullopt;
        };
        const std::optional<int64_t> input_tok = read_i64("input_tokens");
        const std::optional<int64_t> output_tok = read_i64("output_tokens");
        if (input_tok.has_value() || output_tok.has_value()) {
            usage_out->prompt_tokens = input_tok.value_or(0);
            usage_out->output_tokens = output_tok.value_or(0);
            usage_out->total_tokens = usage_out->prompt_tokens + usage_out->output_tokens;
            usage_out->has_usage = true;
        }
    } catch (...) {
    }
}

std::optional<std::string> ExtractAssistantText(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto cit = root.find("content");
        if (cit == root.end() || !cit->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& blocks = cit->value().as_array();
        std::string merged;
        for (const boost::json::value& b : blocks) {
            if (!b.is_object()) {
                continue;
            }
            const boost::json::object& bo = b.as_object();
            const auto tit = bo.find("type");
            if (tit == bo.end() || !tit->value().is_string()) {
                continue;
            }
            if (std::string(tit->value().as_string()) != "text") {
                continue;
            }
            const auto tx = bo.find("text");
            if (tx == bo.end() || !tx->value().is_string()) {
                continue;
            }
            merged.append(std::string(tx->value().as_string()));
        }
        if (merged.empty()) {
            return std::nullopt;
        }
        return merged;
    } catch (...) {
        return std::nullopt;
    }
}

std::string AnthropicExtractPrimaryDocumentJsonImpl(const std::string& api_key_utf8,
                                                    const std::string& model_id_utf8,
                                                    std::string_view mime_type,
                                                    const std::vector<char>& bytes,
                                                    const int max_output_tokens,
                                                    std::string& error_out,
                                                    GeminiUsageStats* usage_out,
                                                    const GeminiHttpTimeouts* timeouts,
                                                    std::string* raw_response_out) {
    error_out.clear();
    if (raw_response_out) {
        raw_response_out->clear();
    }
    if (api_key_utf8.empty()) {
        error_out = "API key is empty";
        return {};
    }
    if (bytes.empty()) {
        error_out =
            (mime_type == "application/pdf") ? "Empty PDF data" : "Empty image data";
        return {};
    }
    if (mime_type == "application/pdf") {
        if (!BytesLookLikePdf(bytes)) {
            error_out = "PDF bytes passed to image method";
            return {};
        }
    } else if (BytesLookLikePdf(bytes)) {
        error_out = "PDF bytes passed to image method";
        return {};
    }
    if (!IsPrimaryInlineMime(mime_type)) {
        error_out = "Invalid inline MIME type";
        return {};
    }
    constexpr size_t kAnthropicMaxRequestBytes = 32U * 1024U * 1024U;
    if (bytes.size() > kAnthropicMaxRequestBytes) {
        error_out = "Anthropic request payload too large";
        return {};
    }

    const int max_tokens =
        std::max(4096, std::min(max_output_tokens > 0 ? max_output_tokens : 32000, 128000));

    std::string model = ResolveAnthropicModelAlias(model_id_utf8);
    if (!IsValidAnthropicModelId(model)) {
        error_out = "Invalid Anthropic model id";
        return {};
    }

    const std::string b64 = Base64EncodeBytes(bytes);
    const std::string user_instruction = GeminiPrimaryDocumentAnthropicUserPromptUtf8();

    boost::json::object source;
    source["type"] = "base64";
    source["media_type"] = std::string(mime_type);
    source["data"] = b64;

    boost::json::object media_block;
    if (mime_type == "application/pdf") {
        media_block["type"] = "document";
        media_block["source"] = source;
    } else {
        media_block["type"] = "image";
        media_block["source"] = source;
    }

    boost::json::object text_block;
    text_block["type"] = "text";
    text_block["text"] = user_instruction;

    boost::json::array content;
    content.push_back(media_block);
    content.push_back(text_block);

    boost::json::object user_msg;
    user_msg["role"] = "user";
    user_msg["content"] = content;

    boost::json::array messages;
    messages.push_back(user_msg);

    boost::json::object root;
    root["model"] = model;
    root["max_tokens"] = max_tokens;
    root["system"] =
        "You are a precise accounting-document data extractor. Output only valid JSON per the user "
        "message; no markdown fences and no commentary outside JSON.";
    root["messages"] = messages;

    const std::string body = boost::json::serialize(root);
    const std::string_view beta =
        (mime_type == "application/pdf") ? std::string_view("pdfs-2024-09-25") : std::string_view();

    unsigned long http_status = 0;
    std::string raw_response;
    constexpr int kMaxApiAttempts = 3;
    for (int attempt = 1; attempt <= kMaxApiAttempts; ++attempt) {
        if (!HttpPostAnthropic(api_key_utf8, body, timeouts, beta, http_status, raw_response, error_out)) {
            if (raw_response_out) {
                *raw_response_out = raw_response;
            }
            return {};
        }
        if (raw_response_out) {
            *raw_response_out = raw_response;
        }
        if (http_status == 200UL) {
            break;
        }
        const auto api_err = ExtractAnthropicApiErrorMessage(raw_response);
        if (IsTransientAnthropicOverload(http_status, api_err) && attempt < kMaxApiAttempts) {
            std::this_thread::sleep_for(std::chrono::milliseconds(500 * attempt));
            continue;
        }
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = std::string("Anthropic API HTTP ") + std::to_string(http_status);
        }
        return {};
    }

    ExtractAnthropicUsage(raw_response, usage_out);

    const auto stop = ExtractAnthropicStopReason(raw_response);
    if (stop.has_value() && *stop == "max_tokens") {
        error_out = "Anthropic output truncated (max_tokens)";
        return {};
    }

    const auto text_opt = ExtractAssistantText(raw_response);
    if (!text_opt.has_value()) {
        const auto api_err = ExtractAnthropicApiErrorMessage(raw_response);
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Anthropic response has no extractable assistant text";
        }
        return {};
    }

    std::string norm_err;
    const std::string out =
        GeminiNormalizePrimaryDocumentJsonFromAssistantText(*text_opt, norm_err);
    if (!norm_err.empty()) {
        error_out = norm_err;
        return {};
    }
    return out;
}

} // namespace

std::string AnthropicGeneratePlainText(const std::string& api_key_utf8,
                                       const std::string& model_id_utf8,
                                       const int max_output_tokens,
                                       const std::string& user_text_utf8,
                                       std::string& error_out,
                                       GeminiUsageStats* usage_out,
                                       const GeminiHttpTimeouts* timeouts,
                                       std::string* raw_response_out) {
    error_out.clear();
    if (raw_response_out) {
        raw_response_out->clear();
    }
    if (api_key_utf8.empty()) {
        error_out = "API key is empty";
        return {};
    }
    if (user_text_utf8.size() > kMaxPlainTextPromptBytes) {
        error_out = "Prompt text too large";
        return {};
    }
    std::string trimmed = user_text_utf8;
    TrimAsciiInPlace(trimmed);
    if (trimmed.empty()) {
        error_out = "Empty prompt text";
        return {};
    }

    const int max_tokens =
        std::max(1, std::min(max_output_tokens > 0 ? max_output_tokens : 8192, 128000));

    std::string model = ResolveAnthropicModelAlias(model_id_utf8);
    if (!IsValidAnthropicModelId(model)) {
        error_out = "Invalid Anthropic model id";
        return {};
    }

    const std::string body = BuildMessagesRequestBody(model, max_tokens, user_text_utf8);

    unsigned long http_status = 0;
    std::string raw_response;
    constexpr int kMaxApiAttempts = 3;
    for (int attempt = 1; attempt <= kMaxApiAttempts; ++attempt) {
        if (!HttpPostAnthropic(api_key_utf8, body, timeouts, {}, http_status, raw_response, error_out)) {
            if (raw_response_out) {
                *raw_response_out = raw_response;
            }
            return {};
        }
        if (raw_response_out) {
            *raw_response_out = raw_response;
        }
        if (http_status == 200UL) {
            break;
        }
        const auto api_err = ExtractAnthropicApiErrorMessage(raw_response);
        if (IsTransientAnthropicOverload(http_status, api_err) && attempt < kMaxApiAttempts) {
            std::this_thread::sleep_for(std::chrono::milliseconds(500 * attempt));
            continue;
        }
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Anthropic API HTTP " + std::to_string(http_status);
        }
        return {};
    }

    ExtractAnthropicUsage(raw_response, usage_out);

    const auto text_opt = ExtractAssistantText(raw_response);
    if (!text_opt.has_value()) {
        const auto api_err = ExtractAnthropicApiErrorMessage(raw_response);
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Anthropic response has no extractable assistant text";
        }
        return {};
    }

    return *text_opt;
}

std::string AnthropicExtractPrimaryDocumentJson(const std::string& api_key_utf8,
                                                 const std::string& model_id_utf8,
                                                 std::string_view mime_type,
                                                 const std::vector<char>& bytes,
                                                 const int max_output_tokens,
                                                 std::string& error_out,
                                                 GeminiUsageStats* usage_out,
                                                 const GeminiHttpTimeouts* timeouts,
                                                 std::string* raw_response_out) {
    return AnthropicExtractPrimaryDocumentJsonImpl(api_key_utf8, model_id_utf8, mime_type, bytes,
                                                   max_output_tokens, error_out, usage_out, timeouts,
                                                   raw_response_out);
}

std::string AnthropicSupportedModelsCatalogJson() {
    boost::json::array models;
    auto add = [&models](const char* id, const char* name, const char* notes_utf8) {
        boost::json::object o;
        o["id"] = id;
        o["name"] = name;
        o["notes"] = std::string(notes_utf8);
        models.push_back(o);
    };

    add("claude-sonnet-4-20250514", "Claude Sonnet 4",
        u8"Сбалансированная линейка 4; идентификатор может обновляться в консоли Anthropic.");
    add("claude-sonnet-4-5-20250929", "Claude Sonnet 4.5",
        u8"Новее Sonnet 4; проверяйте доступность в вашем аккаунте.");
    add("claude-haiku-4-5-20251001", "Claude Haiku 4.5", u8"Быстрее и дешевле; для массовых запросов.");
    add("claude-3-5-sonnet-20241022", "Claude 3.5 Sonnet",
        u8"Стабильный идентификатор линейки 3.5; подходит, если модели 4.x ещё не подключены.");
    add("claude-3-5-haiku-20241022", "Claude 3.5 Haiku", u8"Облегчённая 3.5.");
    add("claude-3-opus-20240229", "Claude 3 Opus", u8"Максимальное качество линейки 3; дороже.");

    add("sonnet45", "Sonnet 4.5 (алиас)", u8"В аддине мапится на claude-sonnet-4-5-20250929.");
    add("haiku45", "Haiku 4.5 (алиас)", u8"В аддине мапится на claude-haiku-4-5-20251001.");
    add("sonnet35", "Sonnet 3.5 (алиас)", u8"В аддине мапится на claude-3-5-sonnet-20241022.");

    boost::json::object root;
    root["defaultModelId"] = std::string(kAnthropicDefaultModelId);
    root["propertyHint"] = std::string(
        u8"В свойство МодельAnthropic (AnthropicModel) подставляйте значение поля id из массива models "
        u8"или краткий алиас (sonnet45, haiku45, sonnet35).");
    root["disclaimer"] = std::string(
        u8"Имена моделей и даты в суффиксах меняются; актуальный список — в консоли Anthropic "
        u8"(https://console.anthropic.com/). Разбор первички (PDF/изображение) при провайдере "
        u8"Anthropic использует Messages API с document/image; выбор провайдера — свойство "
        u8"ПровайдерИИ / AiProvider.");
    root["models"] = models;
    return boost::json::serialize(root);
}
