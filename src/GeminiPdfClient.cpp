#include "GeminiPdfClient.h"

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
#include <cctype>
#include <chrono>
#include <cmath>
#include <codecvt>
#include <iomanip>
#include <cwctype>
#include <limits>
#include <locale>
#include <optional>
#include <sstream>
#include <string>
#include <string_view>
#include <thread>
#include <vector>

namespace {

constexpr char kHostUtf8[] = "generativelanguage.googleapis.com";
#if defined(_WIN32)
constexpr wchar_t kHost[] = L"generativelanguage.googleapis.com";
constexpr wchar_t kPathPrefix[] = L"/v1beta/models/";
constexpr wchar_t kPathSuffix[] = L":generateContent";
#else
constexpr char kPathPrefix[] = "/v1beta/models/";
constexpr char kPathSuffix[] = ":generateContent";
#endif

constexpr int kWinHttpResolveTimeoutMs = 10000;
constexpr int kWinHttpConnectTimeoutMs = 10000;
constexpr int kWinHttpSendTimeoutMs = 30000;
constexpr int kWinHttpReceiveTimeoutMs = 45000;
constexpr int kWinHttpTotalRequestDeadlineMs = 65000;
constexpr int kWinHttpMinReceiveTimeoutMs = 1000;
constexpr int kWinHttpMinTotalDeadlineMs = 5000;

void TrimAsciiInPlace(std::string& s) {
    while (!s.empty() && (static_cast<unsigned char>(s.front()) <= ' ')) {
        s.erase(s.begin());
    }
    while (!s.empty() && (static_cast<unsigned char>(s.back()) <= ' ')) {
        s.pop_back();
    }
}

std::string Base64Encode(const std::vector<char>& in) {
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

/// Допустимые символы в id модели Gemini (исключаем / ? # для пути WinHTTP).
bool IsValidGeminiModelId(std::string_view s) {
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

boost::json::object BuildInvoiceResponseSchema() {
    boost::json::object line_item;
    line_item["type"] = "object";
    {
        boost::json::object props;
        props["name"] = boost::json::object({{"type", "string"}});
        props["sku"] = boost::json::object({{"type", "string"}});
        props["barcode"] = boost::json::object({{"type", "string"}});
        props["quantity"] = boost::json::object({{"type", "string"}});
        props["unit"] = boost::json::object({{"type", "string"}});
        props["price"] = boost::json::object({{"type", "string"}});
        props["priceVatType"] = boost::json::object({{"type", "string"}});
        props["vatRate"] = boost::json::object({{"type", "string"}});
        props["amount"] = boost::json::object({{"type", "string"}});
        line_item["properties"] = props;
    }

    boost::json::object counterparty;
    counterparty["type"] = "object";
    {
        boost::json::object props;
        props["supplier"] = boost::json::object({{"type", "string"}});
        props["buyer"] = boost::json::object({{"type", "string"}});
        props["supplierInn"] = boost::json::object({{"type", "string"}});
        props["buyerInn"] = boost::json::object({{"type", "string"}});
        props["supplierKpp"] = boost::json::object({{"type", "string"}});
        props["buyerKpp"] = boost::json::object({{"type", "string"}});
        props["supplierOkpo"] = boost::json::object({{"type", "string"}});
        props["buyerOkpo"] = boost::json::object({{"type", "string"}});
        counterparty["properties"] = props;
        boost::json::array cp_req;
        cp_req.push_back("supplier");
        cp_req.push_back("buyer");
        cp_req.push_back("supplierInn");
        cp_req.push_back("buyerInn");
        cp_req.push_back("supplierKpp");
        cp_req.push_back("buyerKpp");
        cp_req.push_back("supplierOkpo");
        cp_req.push_back("buyerOkpo");
        counterparty["required"] = cp_req;
    }

    boost::json::object contract;
    contract["type"] = "object";
    {
        boost::json::object props;
        props["numberAndDetails"] = boost::json::object({{"type", "string"}});
        props["date"] = boost::json::object({{"type", "string"}});
        contract["properties"] = props;
    }

    boost::json::object bank;
    bank["type"] = "object";
    {
        boost::json::object props;
        props["bankName"] = boost::json::object({{"type", "string"}});
        props["bik"] = boost::json::object({{"type", "string"}});
        props["correspondentAccount"] = boost::json::object({{"type", "string"}});
        props["settlementAccount"] = boost::json::object({{"type", "string"}});
        bank["properties"] = props;
    }

    boost::json::object line_items;
    line_items["type"] = "array";
    line_items["items"] = line_item;

    boost::json::array doc_type_enum;
    doc_type_enum.push_back("Счет");
    doc_type_enum.push_back("ПриходнаяНакладная");
    doc_type_enum.push_back("РасходнаяНакладная");
    doc_type_enum.push_back("АктВыполненныхРабот");
    doc_type_enum.push_back("Неопределено");
    boost::json::object document_type_field;
    document_type_field["type"] = "string";
    document_type_field["enum"] = doc_type_enum;

    boost::json::object root_props;
    root_props["counterparty"] = counterparty;
    root_props["contract"] = contract;
    root_props["invoiceNumber"] = boost::json::object({{"type", "string"}});
    root_props["documentDate"] = boost::json::object({{"type", "string"}});
    root_props["documentTitle"] = boost::json::object({{"type", "string"}});
    root_props["documentType"] = document_type_field;
    root_props["bankDetails"] = bank;
    root_props["priceColumnVatType"] = boost::json::object({{"type", "string"}});
    root_props["lineItems"] = line_items;
    root_props["rawText"] = boost::json::object({{"type", "string"}});

    boost::json::object schema;
    schema["type"] = "object";
    schema["properties"] = root_props;
    boost::json::array req;
    req.push_back("counterparty");
    req.push_back("contract");
    req.push_back("invoiceNumber");
    req.push_back("documentDate");
    req.push_back("documentTitle");
    req.push_back("documentType");
    req.push_back("bankDetails");
    req.push_back("lineItems");
    schema["required"] = req;
    return schema;
}

bool IsPdfSignature(const std::vector<char>& b) {
    return b.size() >= 5 && b[0] == '%' && b[1] == 'P' && b[2] == 'D' && b[3] == 'F';
}

/// JPEG, PNG, GIF, WebP, BMP, TIFF (без PDF).
std::optional<std::string> DetectRasterImageMime(const std::vector<char>& b) {
    if (b.size() < 2) {
        return std::nullopt;
    }
    const auto* u = reinterpret_cast<const unsigned char*>(b.data());
    const size_t n = b.size();
    if (n >= 3U && u[0] == 0xFFU && u[1] == 0xD8U && u[2] == 0xFFU) {
        return std::string("image/jpeg");
    }
    if (n >= 8U && u[0] == 0x89U && u[1] == 'P' && u[2] == 'N' && u[3] == 'G' && u[4] == 0x0DU
        && u[5] == 0x0AU && u[6] == 0x1AU && u[7] == 0x0AU) {
        return std::string("image/png");
    }
    if (n >= 6U && u[0] == 'G' && u[1] == 'I' && u[2] == 'F' && u[3] == '8'
        && (u[4] == '7' || u[4] == '9') && u[5] == 'a') {
        return std::string("image/gif");
    }
    if (n >= 12U && u[0] == 'R' && u[1] == 'I' && u[2] == 'F' && u[3] == 'F' && u[8] == 'W'
        && u[9] == 'E' && u[10] == 'B' && u[11] == 'P') {
        return std::string("image/webp");
    }
    if (u[0] == 'B' && u[1] == 'M') {
        return std::string("image/bmp");
    }
    if (n >= 4U
        && ((u[0] == 'I' && u[1] == 'I' && u[2] == 0x2AU && u[3] == 0)
            || (u[0] == 'M' && u[1] == 'M' && u[2] == 0 && u[3] == 0x2AU))) {
        return std::string("image/tiff");
    }
    return std::nullopt;
}

bool IsAllowedInlineMime(std::string_view mime) {
    return mime == "application/pdf" || mime == "image/jpeg" || mime == "image/png"
        || mime == "image/gif" || mime == "image/webp" || mime == "image/bmp" || mime == "image/tiff";
}

std::string BuildRequestBody(std::string_view mime_type, const std::string& inline_b64) {
    boost::json::object inline_data;
    inline_data["mime_type"] = std::string(mime_type);
    inline_data["data"] = inline_b64;

    boost::json::object part_binary;
    part_binary["inline_data"] = inline_data;

    const char* prompt =
        "You are given an attached file: either a PDF or a single raster image (e.g. JPEG, PNG). "
        "For PDF: pages may be pure scans, have a selectable text layer, or both — read everything "
        "that is visibly rendered on each page. For one image: read all visible text and layout from "
        "the picture. Before extracting fields, detect page/image orientation and read text in corrected "
        "orientation (including 90 degrees, 180 degrees, and 270 degrees rotations) when needed. The "
        "content may be a primary accounting document (invoice, bill, act, waybill, "
        "receipt, etc.) or another business graphic document; visible text may be Ukrainian, Russian, "
        "English, or other languages shown. Extract all data relevant to the schema and return a single "
        "JSON object that strictly follows the provided JSON schema. Use empty string \"\" for unknown "
        "fields. Preserve the original language and spelling from the document for text values. Field "
        "rawText must contain the full visible text from the whole document/image, not only the title: "
        "include table headers, row text, totals and VAT labels if visible (examples: \"Ціна без ПДВ\", "
        "\"Ціна з ПДВ\", \"Сума ПДВ\", \"Всього без ПДВ\", \"Всього з ПДВ\", \"без НДС\", \"с НДС\"). "
        "Field priceColumnVatType: set \"без НДС\" if table header says price without VAT (e.g. "
        "\"Ціна без ПДВ\", \"Цена без НДС\"); set \"с НДС\" if price with VAT (e.g. \"Ціна з ПДВ\", "
        "\"Цена с НДС\"); else \"\". For "
        "lineItems, include every product/service row with name, sku, barcode, quantity, unit, price, "
        "priceVatType, vatRate, amount when visible. Field sku is the article / vendor code / SKU when the "
        "document shows it (labels like Артикул, Артикул постачальника, Код товару, SKU, Article, "
        "Part number, Cat. no., etc.); use \"\" if no separate article column or value exists. Field "
        "barcode is a product barcode value from labels like Штрихкод, Barcode, EAN, GTIN, UPC "
        "(digits only, keep leading zeros), otherwise \"\". Field priceVatType is \"с НДС\" when the "
        "line price explicitly includes VAT, \"без НДС\" when the line price explicitly excludes VAT; "
        "if VAT status is not explicitly stated for that line but priceColumnVatType is known, copy "
        "priceColumnVatType into each line. Field vatRate is the explicit VAT "
        "rate for that line (examples: \"20%\", \"7%\", \"0%\", \"без НДС\"); if no clear per-line rate is "
        "shown, use \"\". Omit other fields only if absent. "
        "In counterparty: supplierInn is the tax id of the "
        "supplier/seller/issuer (e.g. ЄДРПОУ, ДРФО, ИНН next to Постачальник/Продавець/Виконавець); "
        "buyerInn is the buyer's tax id when shown (Покупець/Замовник). supplierKpp and buyerKpp are "
        "Russian КПП when present; otherwise \"\". supplierOkpo and buyerOkpo: in Russian "
        "documents use label ОКПО (8 digits for organizations, 10 for sole proprietors); in Ukrainian "
        "documents the same role is ЄДРПОУ (typically 8 digits for legal entities) — put that value "
        "in supplierOkpo/buyerOkpo when it is the party's registration/statistical code, even if also "
        "in supplierInn. If absent, \"\". Do not put the buyer's code into supplierOkpo. "
        "For contract.numberAndDetails: the sign № (and prefixes like N/Nr) is not part of the value; "
        "remove it and keep only meaningful contract details/identifier text. "
        "Field documentDate: date of the current primary document itself (invoice/act/waybill date), "
        "not the contract date; output strictly in YYYY/mm/dd format, or \"\" if absent/unclear. "
        "Field documentTitle: the visible printed document title or form name at the top (e.g. "
        "\"Счет на оплату\", \"Приходная накладная\", \"Акт выполненных работ\"); use \"\" if absent. "
        "Field documentType: infer from documentTitle (and header wording) — exactly one enum: "
        "Счет — invoice, bill, payment invoice, счет-фактура, proforma, рахунок (on payment), etc.; "
        "ПриходнаяНакладная — incoming / goods receipt, приходная накладная, прибуткова накладна; "
        "РасходнаяНакладная — outgoing, расходная накладная, видаткова накладна; "
        "АктВыполненныхРабот — act of works/services, акт выполненных работ, акт наданих послуг, "
        "акт прийому-передачі, acceptance certificates for works/services; "
        "Неопределено — when the title does not clearly match any of the above.";

    boost::json::object part_text;
    part_text["text"] = prompt;

    boost::json::array parts;
    parts.push_back(part_binary);
    parts.push_back(part_text);

    boost::json::object content;
    content["parts"] = parts;

    boost::json::array contents;
    contents.push_back(content);

    boost::json::object gen_cfg;
    gen_cfg["responseMimeType"] = "application/json";
    gen_cfg["responseJsonSchema"] = BuildInvoiceResponseSchema();

    boost::json::object root;
    root["contents"] = contents;
    root["generationConfig"] = gen_cfg;

    return std::string(boost::json::serialize(root));
}

constexpr size_t kMaxPlainTextPromptBytes = 2U * 1024U * 1024U;

/// Подсказка модели: не нормализовать пробелы/переносы без явной просьбы пользователя.
constexpr char kPlainTextSystemInstruction[] =
    "When the user message has line breaks, indentation, lists, tables, or multiple paragraphs, treat "
    "that structure as meaningful. In your answer, preserve the same kind of formatting (line breaks, "
    "spacing, bullet/numbered lists, aligned blocks) whenever you reproduce, quote, transform, or "
    "build on their text. Do not collapse whitespace or flatten layout unless the user explicitly "
    "asks for plain or compact output. If the task involves reading text from an image or document "
    "snapshot, first infer orientation and read in corrected orientation (including 90 degrees, "
    "180 degrees, and 270 degrees rotations) when needed.";

std::string BuildPlainTextRequestBody(const std::string& user_text_utf8) {
    boost::json::object sys_part;
    sys_part["text"] = std::string(kPlainTextSystemInstruction);
    boost::json::array sys_parts_arr;
    sys_parts_arr.push_back(sys_part);
    boost::json::object system_instruction;
    system_instruction["parts"] = sys_parts_arr;

    boost::json::object part_text;
    part_text["text"] = user_text_utf8;

    boost::json::array parts;
    parts.push_back(part_text);

    boost::json::object content;
    content["parts"] = parts;

    boost::json::array contents;
    contents.push_back(content);

    boost::json::object root;
    root["systemInstruction"] = system_instruction;
    root["contents"] = contents;

    return std::string(boost::json::serialize(root));
}

bool HttpPostGemini(const std::string& api_key_utf8,
                    const std::string& model_id_utf8,
                    const std::string& body_utf8,
                    const GeminiHttpTimeouts* timeouts,
                    unsigned long& status_out,
                    std::string& response_body_out,
                    std::string& error_out) {
#if defined(_WIN32)
    const auto is_transient_winhttp_error = [](DWORD e) -> bool {
        return e == ERROR_WINHTTP_TIMEOUT || e == ERROR_WINHTTP_NAME_NOT_RESOLVED
            || e == ERROR_WINHTTP_CANNOT_CONNECT || e == ERROR_WINHTTP_CONNECTION_ERROR
            || e == ERROR_WINHTTP_RESEND_REQUEST || e == ERROR_WINHTTP_INTERNAL_ERROR;
    };

    const std::wstring path =
        std::wstring(kPathPrefix) + Utf8ToWide(model_id_utf8) + kPathSuffix;
    const std::wstring wkey = Utf8ToWide(api_key_utf8);
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
        error_out = "Gemini API network error (WinHttpOpen)";
        return false;
    }
    // Fail fast enough for UI usage: keep retries, but cap a single receive wait.
    WinHttpSetTimeouts(session, kWinHttpResolveTimeoutMs, kWinHttpConnectTimeoutMs,
                       kWinHttpSendTimeoutMs, configured_receive_timeout_ms);

    bool ok = false;
    HINTERNET connect =
        WinHttpConnect(session, kHost, INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!connect) {
        error_out = "Gemini API network error (WinHttpConnect)";
        WinHttpCloseHandle(session);
        return false;
    }

    const std::wstring hdr_key = L"x-goog-api-key: " + wkey;
    std::wstring headers = hdr_key + L"\r\nContent-Type: application/json\r\n";
    constexpr int kMaxAttempts = 3;
    const auto started_at = std::chrono::steady_clock::now();
    for (int attempt = 1; attempt <= kMaxAttempts; ++attempt) {
        const auto now = std::chrono::steady_clock::now();
        const auto elapsed_ms =
            std::chrono::duration_cast<std::chrono::milliseconds>(now - started_at).count();
        const int remaining_total_ms =
            configured_total_deadline_ms - static_cast<int>(elapsed_ms);
        if (remaining_total_ms <= 0) {
            error_out = "Gemini API timeout: total request deadline exceeded";
            break;
        }
        const int receive_timeout_ms = std::min(configured_receive_timeout_ms, remaining_total_ms);

        HINTERNET request =
            WinHttpOpenRequest(connect, L"POST", path.c_str(), nullptr, WINHTTP_NO_REFERER,
                               WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
        if (!request) {
            const DWORD e = GetLastError();
            error_out = "Gemini API network error (WinHttpOpenRequest, code=" + std::to_string(e) + ")";
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
            error_out = "Gemini API network error (WinHttpSendRequest, code=" + std::to_string(e) + ")";
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
                error_out = "Gemini API timeout in WinHttpReceiveResponse (code=12002)";
            } else {
                error_out = "Gemini API network error (WinHttpReceiveResponse, code=" + std::to_string(e) + ")";
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
                error_out = "Gemini API network error (WinHttpQueryDataAvailable, code="
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
                error_out = "Gemini API network error (WinHttpReadData, code=" + std::to_string(e) + ")";
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
    return ok;
#else
    namespace beast = boost::beast;
    namespace http = beast::http;
    namespace net = boost::asio;
    namespace ssl = boost::asio::ssl;
    using tcp = net::ip::tcp;

    try {
        const std::string path = std::string(kPathPrefix) + model_id_utf8 + kPathSuffix;
        net::io_context ioc;
        ssl::context ctx(ssl::context::tls_client);
        ctx.set_default_verify_paths();

        tcp::resolver resolver(ioc);
        beast::ssl_stream<beast::tcp_stream> stream(ioc, ctx);
        if (!SSL_set_tlsext_host_name(stream.native_handle(), kHostUtf8)) {
            error_out = "Gemini API TLS error (SNI)";
            return false;
        }

        const auto results = resolver.resolve(kHostUtf8, "443");
        beast::get_lowest_layer(stream).connect(results);
        stream.handshake(ssl::stream_base::client);

        http::request<http::string_body> req{http::verb::post, path, 11};
        req.set(http::field::host, kHostUtf8);
        req.set(http::field::user_agent, "VisualRecognitionAddIn/1.0");
        req.set("x-goog-api-key", api_key_utf8);
        req.set(http::field::content_type, "application/json");
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
            error_out = "Gemini API TLS shutdown error";
            return false;
        }
        return true;
    } catch (const std::exception& e) {
        error_out = std::string("Gemini API network error: ") + e.what();
        return false;
    }
#endif
}

std::optional<std::string> ExtractGeminiApiErrorMessage(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& o = v.as_object();
        const auto it = o.find("error");
        if (it == o.end() || !it->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& err = it->value().as_object();
        const auto im = err.find("message");
        if (im == err.end() || !im->value().is_string()) {
            return std::string("Gemini API error");
        }
        return std::string(im->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<std::string> ExtractPromptBlockMessage(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto pc = root.find("promptFeedback");
        if (pc == root.end() || !pc->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& pf = pc->value().as_object();
        const auto pb = pf.find("blockReason");
        if (pb == pf.end()) {
            return std::nullopt;
        }
        if (pb->value().is_string()) {
            return std::string(pb->value().as_string());
        }
        return std::string("blocked");
    } catch (...) {
        return std::nullopt;
    }
}

bool IsTransientGeminiOverload(unsigned long http_status, const std::optional<std::string>& api_err) {
    if (http_status == 429UL || http_status == 503UL) {
        return true;
    }
    if (!api_err.has_value()) {
        return false;
    }
    std::string s = *api_err;
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    return s.find("high demand") != std::string::npos || s.find("try again later") != std::string::npos
        || s.find("resource exhausted") != std::string::npos || s.find("temporar") != std::string::npos;
}

std::optional<std::string> ExtractCandidateText(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto it = root.find("candidates");
        if (it == root.end() || !it->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& cand = it->value().as_array();
        if (cand.empty()) {
            return std::nullopt;
        }
        const boost::json::value& c0 = cand[0];
        if (!c0.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& cobj = c0.as_object();
        const auto cit = cobj.find("content");
        if (cit == cobj.end() || !cit->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& content = cit->value().as_object();
        const auto pit = content.find("parts");
        if (pit == content.end() || !pit->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& parts = pit->value().as_array();
        if (parts.empty() || !parts[0].is_object()) {
            return std::nullopt;
        }
        const boost::json::object& p0 = parts[0].as_object();
        const auto tit = p0.find("text");
        if (tit == p0.end() || !tit->value().is_string()) {
            return std::nullopt;
        }
        return std::string(tit->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}

/// Все текстовые части первого кандидата подряд (модель может разбить длинный ответ на несколько parts).
std::optional<std::string> ExtractCandidateTextAllTextParts(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto it = root.find("candidates");
        if (it == root.end() || !it->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& cand = it->value().as_array();
        if (cand.empty()) {
            return std::nullopt;
        }
        const boost::json::value& c0 = cand[0];
        if (!c0.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& cobj = c0.as_object();
        const auto cit = cobj.find("content");
        if (cit == cobj.end() || !cit->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& content = cit->value().as_object();
        const auto pit = content.find("parts");
        if (pit == content.end() || !pit->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& parts = pit->value().as_array();
        std::string merged;
        for (const boost::json::value& pv : parts) {
            if (!pv.is_object()) {
                continue;
            }
            const boost::json::object& po = pv.as_object();
            const auto tit = po.find("text");
            if (tit == po.end() || !tit->value().is_string()) {
                continue;
            }
            merged.append(std::string(tit->value().as_string()));
        }
        if (merged.empty()) {
            return std::nullopt;
        }
        return merged;
    } catch (...) {
        return std::nullopt;
    }
}

void ExtractUsageStats(const std::string& json, GeminiUsageStats* usage_out) {
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
        const auto uit = root.find("usageMetadata");
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
                const uint64_t v = it->value().as_uint64();
                if (v <= static_cast<uint64_t>(std::numeric_limits<int64_t>::max())) {
                    return static_cast<int64_t>(v);
                }
            }
            return std::nullopt;
        };

        const std::optional<int64_t> prompt = read_i64("promptTokenCount");
        const std::optional<int64_t> output = read_i64("candidatesTokenCount");
        const std::optional<int64_t> total = read_i64("totalTokenCount");
        if (prompt.has_value() || output.has_value() || total.has_value()) {
            usage_out->prompt_tokens = prompt.value_or(0);
            usage_out->output_tokens = output.value_or(0);
            usage_out->total_tokens = total.value_or(0);
            usage_out->has_usage = true;
        }
    } catch (...) {
    }
}

std::string StripMarkdownJsonFence(std::string s) {
    TrimAsciiInPlace(s);
    if (s.size() >= 7 && s.compare(0, 7, "```json") == 0) {
        s.erase(0, 7);
        TrimAsciiInPlace(s);
    } else if (s.size() >= 3 && s.compare(0, 3, "```") == 0) {
        s.erase(0, 3);
        TrimAsciiInPlace(s);
    }
    if (s.size() >= 3 && s.compare(s.size() - 3, 3, "```") == 0) {
        s.resize(s.size() - 3);
        TrimAsciiInPlace(s);
    }
    return s;
}

std::string JsonObjectGetString(const boost::json::object& o, const char* key) {
    const auto it = o.find(key);
    if (it == o.end() || !it->value().is_string()) {
        return {};
    }
    return std::string(it->value().as_string());
}

bool WideContains(const std::wstring& w, const wchar_t* sub) {
    return w.find(sub) != std::wstring::npos;
}

std::wstring CanonicalizeForMatch(std::wstring w) {
    std::transform(w.begin(), w.end(), w.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });
    std::wstring out;
    out.reserve(w.size());
    for (wchar_t ch : w) {
        if (std::iswalnum(ch) != 0) {
            out.push_back(ch);
        }
    }
    return out;
}

std::wstring CanonicalizeUtf8ForMatch(std::string_view text_utf8) {
    return CanonicalizeForMatch(Utf8ToWide(std::string(text_utf8)));
}

bool ContainsZeroVatTotalHint(std::string_view text_utf8) {
    std::wstring w = Utf8ToWide(std::string(text_utf8));
    if (w.empty()) {
        return false;
    }
    std::transform(w.begin(), w.end(), w.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });
    const std::wstring wc = CanonicalizeForMatch(w);

    const bool has_vat_total_label =
        WideContains(w, L"сума пдв") || WideContains(w, L"сумма ндс")
        || WideContains(w, L"итого ндс") || WideContains(w, L"пдв")
        || wc.find(L"сумапдв") != std::wstring::npos || wc.find(L"суммандс") != std::wstring::npos
        || wc.find(L"итогондс") != std::wstring::npos;
    if (!has_vat_total_label) {
        return false;
    }

    const bool has_zero_amount =
        WideContains(w, L"0,00") || WideContains(w, L"0.00") || WideContains(w, L" 0 ")
        || WideContains(w, L":0") || WideContains(w, L"=0");
    return has_zero_amount;
}

std::optional<double> ParseLocalizedNumberToken(const std::wstring& token) {
    std::string ascii;
    ascii.reserve(token.size());
    bool has_digit = false;
    for (wchar_t ch : token) {
        if (ch >= L'0' && ch <= L'9') {
            ascii.push_back(static_cast<char>(ch));
            has_digit = true;
            continue;
        }
        if (ch == L',' || ch == L'.') {
            ascii.push_back('.');
            continue;
        }
        if (ch == L' ' || ch == 0x00A0) {
            continue;
        }
        if (has_digit) {
            break;
        }
    }
    if (!has_digit || ascii.empty()) {
        return std::nullopt;
    }
    try {
        size_t used = 0;
        const double v = std::stod(ascii, &used);
        if (used == 0 || v < 0.0) {
            return std::nullopt;
        }
        return v;
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<double> ExtractNumberAfterAnyLabel(const std::wstring& lowered_text,
                                                 const std::vector<std::wstring>& labels) {
    for (const std::wstring& label : labels) {
        const size_t pos = lowered_text.find(label);
        if (pos == std::wstring::npos) {
            continue;
        }
        const size_t start = pos + label.size();
        const size_t max_len = std::min<size_t>(80, lowered_text.size() - start);
        const std::wstring window = lowered_text.substr(start, max_len);
        const auto parsed = ParseLocalizedNumberToken(window);
        if (parsed.has_value()) {
            return parsed;
        }
    }
    return std::nullopt;
}

std::string NormalizePercentRate(double rate) {
    static const int kCommonRates[] = {0, 7, 10, 14, 20};
    for (int r : kCommonRates) {
        if (std::abs(rate - static_cast<double>(r)) <= 0.35) {
            return std::to_string(r) + "%";
        }
    }
    std::ostringstream oss;
    oss.setf(std::ios::fixed);
    oss.precision(2);
    oss << rate;
    std::string out = oss.str();
    while (!out.empty() && out.back() == '0') {
        out.pop_back();
    }
    if (!out.empty() && out.back() == '.') {
        out.pop_back();
    }
    return out + "%";
}

std::string InferVatRateFromDocumentTotals(std::string_view text_utf8) {
    std::wstring w = Utf8ToWide(std::string(text_utf8));
    if (w.empty()) {
        return {};
    }
    std::transform(w.begin(), w.end(), w.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });

    const auto vat_sum = ExtractNumberAfterAnyLabel(
        w, {L"сума пдв", L"сумма ндс", L"итого ндс", L"у т.ч. пдв", L"в т.ч. ндс", L"ндс:", L"пдв:"});
    const auto base_sum = ExtractNumberAfterAnyLabel(
        w, {L"всього без пдв", L"разом без пдв", L"всего без ндс", L"итого без ндс"});
    const auto total_sum = ExtractNumberAfterAnyLabel(
        w, {L"всього з пдв", L"разом з пдв", L"всего с ндс", L"итого с ндс", L"всього:", L"разом:", L"всего:"});

    if (vat_sum.has_value() && vat_sum.value() == 0.0) {
        return "0%";
    }
    if (vat_sum.has_value() && base_sum.has_value() && base_sum.value() > 0.0) {
        const double rate = (vat_sum.value() / base_sum.value()) * 100.0;
        if (rate >= 0.0 && rate <= 100.0) {
            return NormalizePercentRate(rate);
        }
    }
    if (total_sum.has_value() && base_sum.has_value() && base_sum.value() > 0.0
        && total_sum.value() >= base_sum.value()) {
        const double rate = ((total_sum.value() - base_sum.value()) / base_sum.value()) * 100.0;
        if (rate >= 0.0 && rate <= 100.0) {
            return NormalizePercentRate(rate);
        }
    }
    return {};
}

/// Подбор типа по тексту названия (дополнение к ответу модели).
std::string InferDocumentTypeFromTitle(std::string_view title_utf8) {
    if (title_utf8.empty()) {
        return {};
    }
    std::wstring w = Utf8ToWide(std::string(title_utf8));
    if (w.empty()) {
        return {};
    }
    std::transform(w.begin(), w.end(), w.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });

    const bool incoming_ru = WideContains(w, L"приход") && WideContains(w, L"наклад");
    const bool incoming_ua = WideContains(w, L"прибут") && WideContains(w, L"наклад");
    const bool outgoing_ru = WideContains(w, L"расход") && WideContains(w, L"наклад");
    const bool outgoing_ua = WideContains(w, L"видат") && WideContains(w, L"наклад");

    if (incoming_ru || incoming_ua) {
        return "ПриходнаяНакладная";
    }
    if (outgoing_ru || outgoing_ua) {
        return "РасходнаяНакладная";
    }

    if (WideContains(w, L"акт")) {
        if (WideContains(w, L"выполнен") || WideContains(w, L"виконан") || WideContains(w, L"робіт")
            || WideContains(w, L"послуг") || WideContains(w, L"прийом") || WideContains(w, L"передач")
            || WideContains(w, L"наданих") || WideContains(w, L"здачі") || WideContains(w, L"приймання")) {
            return "АктВыполненныхРабот";
        }
    }

    if (WideContains(w, L"счет") || WideContains(w, L"счёт") || WideContains(w, L"invoice")
        || WideContains(w, L"рахунк") || WideContains(w, L"proforma") || WideContains(w, L"фактура")) {
        return "Счет";
    }

    return {};
}

void MaybeEnrichDocumentType(boost::json::object& root) {
    const std::string dtype = JsonObjectGetString(root, "documentType");
    const std::string title = JsonObjectGetString(root, "documentTitle");
    const bool need_infer = dtype.empty() || dtype == "Неопределено";
    if (!need_infer) {
        return;
    }
    const std::string inferred = InferDocumentTypeFromTitle(title);
    if (!inferred.empty()) {
        root["documentType"] = inferred;
    }
}

std::string NormalizeContractNumberAndDetailsText(std::string_view in_utf8) {
    std::string s(in_utf8);
    TrimAsciiInPlace(s);
    if (s.empty()) {
        return {};
    }

    std::wstring w = Utf8ToWide(s);
    if (w.empty()) {
        return s;
    }

    std::wstring out;
    out.reserve(w.size());
    for (wchar_t ch : w) {
        if (ch == L'№') {
            continue;
        }
        out.push_back(ch);
    }

    // Remove leading "N", "N." or "Nr" prefixes after optional spaces.
    size_t i = 0;
    while (i < out.size() && std::iswspace(out[i]) != 0) {
        ++i;
    }
    if (i < out.size() && (out[i] == L'N' || out[i] == L'n')) {
        size_t j = i + 1;
        if (j < out.size() && (out[j] == L'r' || out[j] == L'R')) {
            ++j;
        }
        if (j < out.size() && out[j] == L'.') {
            ++j;
        }
        while (j < out.size() && std::iswspace(out[j]) != 0) {
            ++j;
        }
        out.erase(0, j);
    }

    std::string normalized;
#if defined(_WIN32)
    if (!out.empty()) {
        const int n = WideCharToMultiByte(CP_UTF8, 0, out.data(), static_cast<int>(out.size()),
                                          nullptr, 0, nullptr, nullptr);
        if (n > 0) {
            normalized.resize(static_cast<size_t>(n));
            WideCharToMultiByte(CP_UTF8, 0, out.data(), static_cast<int>(out.size()),
                                normalized.data(), n, nullptr, nullptr);
        }
    }
#else
    static std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> cvt_utf8_utf16;
    normalized = cvt_utf8_utf16.to_bytes(out.data(), out.data() + out.size());
#endif
    TrimAsciiInPlace(normalized);
    return normalized.empty() ? s : normalized;
}

bool IsLeapYear(int year) {
    return (year % 400 == 0) || ((year % 4 == 0) && (year % 100 != 0));
}

int DaysInMonth(int year, int month) {
    if (month < 1 || month > 12) {
        return 0;
    }
    static const int kDays[] = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};
    if (month == 2) {
        return IsLeapYear(year) ? 29 : 28;
    }
    return kDays[month - 1];
}

bool IsValidYmd(int year, int month, int day) {
    if (year < 1900 || year > 2100) {
        return false;
    }
    const int dim = DaysInMonth(year, month);
    return dim > 0 && day >= 1 && day <= dim;
}

std::string FormatYmdSlash(int year, int month, int day) {
    std::ostringstream oss;
    oss << std::setfill('0') << std::setw(4) << year << "/" << std::setw(2) << month << "/"
        << std::setw(2) << day;
    return oss.str();
}

std::string NormalizeDocumentDateToYmdSlash(std::string_view value_utf8) {
    std::string s(value_utf8);
    TrimAsciiInPlace(s);
    if (s.empty()) {
        return {};
    }

    if (s.size() == 8U && std::all_of(s.begin(), s.end(), [](unsigned char ch) { return std::isdigit(ch) != 0; })) {
        const int year = std::stoi(s.substr(0, 4));
        const int month = std::stoi(s.substr(4, 2));
        const int day = std::stoi(s.substr(6, 2));
        if (IsValidYmd(year, month, day)) {
            return FormatYmdSlash(year, month, day);
        }
    }

    std::vector<int> nums;
    nums.reserve(8);
    int current = -1;
    for (char ch : s) {
        const unsigned char uc = static_cast<unsigned char>(ch);
        if (std::isdigit(uc) != 0) {
            const int digit = static_cast<int>(ch - '0');
            current = (current < 0) ? digit : (current * 10 + digit);
        } else if (current >= 0) {
            nums.push_back(current);
            current = -1;
        }
    }
    if (current >= 0) {
        nums.push_back(current);
    }
    if (nums.size() < 3U) {
        return {};
    }

    for (size_t i = 0; i + 2U < nums.size(); ++i) {
        const int a = nums[i];
        const int b = nums[i + 1U];
        const int c = nums[i + 2U];
        if (IsValidYmd(a, b, c)) {
            return FormatYmdSlash(a, b, c); // YYYY/MM/DD
        }
        if (IsValidYmd(c, b, a)) {
            return FormatYmdSlash(c, b, a); // DD/MM/YYYY
        }
    }
    return {};
}

void MaybeNormalizeDocumentDateField(boost::json::object& root) {
    auto it = root.find("documentDate");
    if (it == root.end()) {
        return;
    }
    const std::string raw = JsonObjectGetString(root, "documentDate");
    if (raw.empty()) {
        return;
    }
    root["documentDate"] = NormalizeDocumentDateToYmdSlash(raw);
}

void MaybeNormalizeContractFields(boost::json::object& root) {
    auto it = root.find("contract");
    if (it == root.end() || !it->value().is_object()) {
        return;
    }
    boost::json::object& contract = it->value().as_object();
    const std::string raw_number = JsonObjectGetString(contract, "numberAndDetails");
    if (!raw_number.empty()) {
        contract["numberAndDetails"] = NormalizeContractNumberAndDetailsText(raw_number);
    }

    const std::string raw_date = JsonObjectGetString(contract, "date");
    if (!raw_date.empty()) {
        contract["date"] = NormalizeDocumentDateToYmdSlash(raw_date);
    }
}

std::string NormalizeVatRateText(std::string_view value_utf8) {
    std::string s(value_utf8);
    TrimAsciiInPlace(s);
    if (s.empty()) {
        return {};
    }

    std::wstring w = Utf8ToWide(s);
    if (!w.empty()) {
        std::transform(w.begin(), w.end(), w.begin(),
                       [](wchar_t ch) { return std::towlower(ch); });
        if (WideContains(w, L"без ндс") || WideContains(w, L"без пдв") || WideContains(w, L"no vat")
            || WideContains(w, L"vat exempt")) {
            return "без НДС";
        }
    }

    std::string number;
    bool seen_digit = false;
    for (char ch : s) {
        const unsigned char uc = static_cast<unsigned char>(ch);
        if (std::isdigit(uc) != 0) {
            number.push_back(ch);
            seen_digit = true;
            continue;
        }
        if (seen_digit && (ch == '.' || ch == ',')) {
            number.push_back('.');
            continue;
        }
        if (seen_digit) {
            break;
        }
    }
    while (!number.empty() && number.back() == '.') {
        number.pop_back();
    }
    if (number.empty()) {
        return {};
    }
    return number + "%";
}

std::string NormalizePriceVatTypeText(std::string_view value_utf8, std::string_view normalized_rate_utf8) {
    std::string s(value_utf8);
    TrimAsciiInPlace(s);
    if (!s.empty()) {
        std::wstring w = Utf8ToWide(s);
        const std::wstring wc = CanonicalizeForMatch(w);
        if (!w.empty()) {
            std::transform(w.begin(), w.end(), w.begin(),
                           [](wchar_t ch) { return std::towlower(ch); });
            if (WideContains(w, L"без ндс") || WideContains(w, L"без пдв") || WideContains(w, L"no vat")
                || WideContains(w, L"without vat") || wc.find(L"безндс") != std::wstring::npos
                || wc.find(L"безпдв") != std::wstring::npos || wc.find(L"novat") != std::wstring::npos
                || wc.find(L"withoutvat") != std::wstring::npos) {
                return "без НДС";
            }
            if (WideContains(w, L"с ндс") || WideContains(w, L"в т.ч. ндс") || WideContains(w, L"з пдв")
                || WideContains(w, L"with vat") || WideContains(w, L"incl")
                || wc.find(L"сндс") != std::wstring::npos || wc.find(L"втчндс") != std::wstring::npos
                || wc.find(L"зпдв") != std::wstring::npos || wc.find(L"withvat") != std::wstring::npos
                || wc.find(L"incl") != std::wstring::npos) {
                return "с НДС";
            }
        }
    }

    std::string rate(normalized_rate_utf8);
    TrimAsciiInPlace(rate);
    if (rate == "без НДС") {
        return "без НДС";
    }
    if (!rate.empty()) {
        return "с НДС";
    }
    return {};
}

std::string InferGlobalPriceVatTypeFromRoot(const boost::json::object& root) {
    const std::string from_column = NormalizePriceVatTypeText(
        JsonObjectGetString(root, "priceColumnVatType"), std::string_view{});
    if (!from_column.empty()) {
        return from_column;
    }

    std::string hints = JsonObjectGetString(root, "rawText");
    const std::string title = JsonObjectGetString(root, "documentTitle");
    if (!title.empty()) {
        if (!hints.empty()) {
            hints.push_back(' ');
        }
        hints += title;
    }
    TrimAsciiInPlace(hints);
    if (hints.empty()) {
        return {};
    }

    std::wstring w = Utf8ToWide(hints);
    if (w.empty()) {
        return {};
    }
    std::transform(w.begin(), w.end(), w.begin(),
                   [](wchar_t ch) { return std::towlower(ch); });
    const std::wstring wc = CanonicalizeForMatch(w);

    const bool without_vat =
        WideContains(w, L"ціна без пдв") || WideContains(w, L"цiна без пдв")
        || WideContains(w, L"цена без ндс") || WideContains(w, L"без пдв")
        || WideContains(w, L"без ндс") || WideContains(w, L"без пдв.")
        || WideContains(w, L"без ндс.") || wc.find(L"цінабезпдв") != std::wstring::npos
        || wc.find(L"цiнабезпдв") != std::wstring::npos || wc.find(L"ценабезндс") != std::wstring::npos
        || wc.find(L"безпдв") != std::wstring::npos || wc.find(L"безндс") != std::wstring::npos;
    if (without_vat) {
        return "без НДС";
    }

    const bool with_vat =
        WideContains(w, L"ціна з пдв") || WideContains(w, L"цiна з пдв")
        || WideContains(w, L"цена с ндс") || WideContains(w, L"с ндс")
        || WideContains(w, L"з пдв") || WideContains(w, L"в т.ч. ндс")
        || WideContains(w, L"у т.ч. пдв") || wc.find(L"ціназпдв") != std::wstring::npos
        || wc.find(L"цiназпдв") != std::wstring::npos || wc.find(L"ценасндс") != std::wstring::npos
        || wc.find(L"сндс") != std::wstring::npos || wc.find(L"зпдв") != std::wstring::npos;
    if (with_vat) {
        return "с НДС";
    }
    return {};
}

void MaybeNormalizeLineItemsVat(boost::json::object& root) {
    auto it = root.find("lineItems");
    if (it == root.end() || !it->value().is_array()) {
        return;
    }
    std::string raw_text = JsonObjectGetString(root, "rawText");
    const std::string doc_title = JsonObjectGetString(root, "documentTitle");
    if (!doc_title.empty()) {
        if (!raw_text.empty()) {
            raw_text.push_back(' ');
        }
        raw_text += doc_title;
    }
    const bool zero_vat_total = ContainsZeroVatTotalHint(raw_text);
    const std::string global_vat_type = InferGlobalPriceVatTypeFromRoot(root);
    const std::string inferred_doc_vat_rate = InferVatRateFromDocumentTotals(raw_text);
    boost::json::array& items = it->value().as_array();
    for (boost::json::value& item_v : items) {
        if (!item_v.is_object()) {
            continue;
        }
        boost::json::object& item = item_v.as_object();
        const std::string vat_rate_raw = JsonObjectGetString(item, "vatRate");
        const std::string vat_type_raw = JsonObjectGetString(item, "priceVatType");

        std::string vat_rate_norm = NormalizeVatRateText(vat_rate_raw);
        std::string vat_type_norm = NormalizePriceVatTypeText(vat_type_raw, vat_rate_norm);
        const std::string column_vat_type = NormalizePriceVatTypeText(
            JsonObjectGetString(root, "priceColumnVatType"), std::string_view{});

        if (vat_rate_norm.empty() && vat_type_norm == "без НДС") {
            vat_rate_norm = "без НДС";
        }
        if (vat_type_norm.empty() && vat_rate_norm == "без НДС") {
            vat_type_norm = "без НДС";
        }
        if (vat_type_norm.empty() && !vat_rate_norm.empty() && vat_rate_norm != "без НДС") {
            vat_type_norm = "с НДС";
        }
        if (!column_vat_type.empty()) {
            vat_type_norm = column_vat_type;
        }
        if (vat_type_norm.empty() && !global_vat_type.empty()) {
            vat_type_norm = global_vat_type;
        }
        if (vat_type_norm.empty() && zero_vat_total) {
            vat_type_norm = "без НДС";
        }
        if (vat_rate_norm.empty() && !inferred_doc_vat_rate.empty()) {
            vat_rate_norm = inferred_doc_vat_rate;
        }
        if (vat_rate_norm.empty() && zero_vat_total) {
            vat_rate_norm = "0%";
        }

        if (!vat_rate_norm.empty() || !vat_rate_raw.empty()) {
            item["vatRate"] = vat_rate_norm;
        }
        if (!vat_type_norm.empty() || !vat_type_raw.empty()) {
            item["priceVatType"] = vat_type_norm;
        }
    }
}

void ForceCopyPriceColumnVatTypeToLines(boost::json::object& root) {
    const std::string raw_column_vat_type = JsonObjectGetString(root, "priceColumnVatType");
    std::string trimmed = raw_column_vat_type;
    TrimAsciiInPlace(trimmed);
    if (trimmed.empty()) {
        return;
    }
    auto it = root.find("lineItems");
    if (it == root.end() || !it->value().is_array()) {
        return;
    }
    boost::json::array& items = it->value().as_array();
    for (boost::json::value& item_v : items) {
        if (!item_v.is_object()) {
            continue;
        }
        boost::json::object& item = item_v.as_object();
        item["priceVatType"] = raw_column_vat_type;
    }
}

std::string GeminiExtractPrimaryDocumentJsonImpl(const std::string& api_key_utf8,
                                                 const std::string& model_id_utf8,
                                                 std::string_view mime_type,
                                                 const std::vector<char>& bytes,
                                                 std::string& error_out,
                                                 GeminiUsageStats* usage_out,
                                                 const GeminiHttpTimeouts* timeouts) {
    error_out.clear();
    if (api_key_utf8.empty()) {
        error_out = "API key is empty";
        return {};
    }
    if (bytes.empty()) {
        error_out =
            (mime_type == "application/pdf") ? "Empty PDF data" : "Empty image data";
        return {};
    }
    if (!IsAllowedInlineMime(mime_type)) {
        error_out = "Invalid inline MIME type";
        return {};
    }

    std::string model = model_id_utf8;
    TrimAsciiInPlace(model);
    if (model.empty()) {
        model = kGeminiDefaultModelId;
    }
    if (!IsValidGeminiModelId(model)) {
        error_out = "Invalid Gemini model id";
        return {};
    }

    const std::string b64 = Base64Encode(bytes);
    const std::string body = BuildRequestBody(mime_type, b64);

    unsigned long http_status = 0;
    std::string raw_response;
    constexpr int kMaxApiAttempts = 3;
    for (int attempt = 1; attempt <= kMaxApiAttempts; ++attempt) {
        if (!HttpPostGemini(api_key_utf8, model, body, timeouts, http_status, raw_response, error_out)) {
            return {};
        }
        if (http_status == 200UL) {
            break;
        }
        const auto api_err = ExtractGeminiApiErrorMessage(raw_response);
        if (IsTransientGeminiOverload(http_status, api_err) && attempt < kMaxApiAttempts) {
            std::this_thread::sleep_for(std::chrono::milliseconds(500 * attempt));
            continue;
        }
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Gemini API HTTP " + std::to_string(http_status);
        }
        return {};
    }

    ExtractUsageStats(raw_response, usage_out);

    const auto text_opt = ExtractCandidateText(raw_response);
    if (!text_opt.has_value()) {
        const auto block = ExtractPromptBlockMessage(raw_response);
        if (block.has_value()) {
            error_out = "Gemini blocked request: " + *block;
        } else {
            error_out = "Gemini response has no extractable text (empty candidates)";
        }
        return {};
    }

    std::string inner = StripMarkdownJsonFence(*text_opt);
    TrimAsciiInPlace(inner);
    if (inner.empty()) {
        error_out = "Gemini returned empty JSON body";
        return {};
    }

    try {
        boost::json::value parsed = boost::json::parse(inner);
        if (parsed.is_object()) {
            MaybeEnrichDocumentType(parsed.as_object());
            MaybeNormalizeDocumentDateField(parsed.as_object());
            MaybeNormalizeContractFields(parsed.as_object());
            MaybeNormalizeLineItemsVat(parsed.as_object());
            ForceCopyPriceColumnVatTypeToLines(parsed.as_object());
        }
        return boost::json::serialize(parsed);
    } catch (const std::exception& e) {
        error_out = std::string("Gemini JSON parse error: ") + e.what();
        return {};
    }
}

std::string GeminiGeneratePlainTextImpl(const std::string& api_key_utf8,
                                          const std::string& model_id_utf8,
                                          const std::string& user_text_utf8,
                                          std::string& error_out,
                                          GeminiUsageStats* usage_out,
                                          const GeminiHttpTimeouts* timeouts) {
    error_out.clear();
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

    std::string model = model_id_utf8;
    TrimAsciiInPlace(model);
    if (model.empty()) {
        model = kGeminiDefaultModelId;
    }
    if (!IsValidGeminiModelId(model)) {
        error_out = "Invalid Gemini model id";
        return {};
    }

    const std::string body = BuildPlainTextRequestBody(user_text_utf8);

    unsigned long http_status = 0;
    std::string raw_response;
    constexpr int kMaxApiAttempts = 3;
    for (int attempt = 1; attempt <= kMaxApiAttempts; ++attempt) {
        if (!HttpPostGemini(api_key_utf8, model, body, timeouts, http_status, raw_response, error_out)) {
            return {};
        }
        if (http_status == 200UL) {
            break;
        }
        const auto api_err = ExtractGeminiApiErrorMessage(raw_response);
        if (IsTransientGeminiOverload(http_status, api_err) && attempt < kMaxApiAttempts) {
            std::this_thread::sleep_for(std::chrono::milliseconds(500 * attempt));
            continue;
        }
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Gemini API HTTP " + std::to_string(http_status);
        }
        return {};
    }

    ExtractUsageStats(raw_response, usage_out);

    const auto text_opt = ExtractCandidateTextAllTextParts(raw_response);
    if (!text_opt.has_value()) {
        const auto block = ExtractPromptBlockMessage(raw_response);
        if (block.has_value()) {
            error_out = "Gemini blocked request: " + *block;
        } else {
            error_out = "Gemini response has no extractable text (empty candidates)";
        }
        return {};
    }

    return *text_opt;
}

} // namespace

std::string GeminiExtractPrimaryDocumentJson(const std::string& api_key_utf8,
                                             const std::string& model_id_utf8,
                                             const std::vector<char>& pdf_bytes,
                                             std::string& error_out,
                                             GeminiUsageStats* usage_out,
                                             const GeminiHttpTimeouts* timeouts) {
    return GeminiExtractPrimaryDocumentJsonImpl(api_key_utf8, model_id_utf8, "application/pdf",
                                                pdf_bytes, error_out, usage_out, timeouts);
}

std::string GeminiExtractPrimaryDocumentJsonFromImageBytes(const std::string& api_key_utf8,
                                                           const std::string& model_id_utf8,
                                                           const std::vector<char>& image_bytes,
                                                           std::string& error_out,
                                                           GeminiUsageStats* usage_out,
                                                           const GeminiHttpTimeouts* timeouts) {
    error_out.clear();
    if (api_key_utf8.empty()) {
        error_out = "API key is empty";
        return {};
    }
    if (image_bytes.empty()) {
        error_out = "Empty image data";
        return {};
    }
    if (IsPdfSignature(image_bytes)) {
        error_out = "PDF bytes passed to image method";
        return {};
    }
    const std::optional<std::string> mime = DetectRasterImageMime(image_bytes);
    if (!mime.has_value()) {
        error_out = "Unsupported inline image format";
        return {};
    }
    return GeminiExtractPrimaryDocumentJsonImpl(api_key_utf8, model_id_utf8, *mime, image_bytes,
                                                error_out, usage_out, timeouts);
}

std::string GeminiGeneratePlainText(const std::string& api_key_utf8,
                                    const std::string& model_id_utf8,
                                    const std::string& user_text_utf8,
                                    std::string& error_out,
                                    GeminiUsageStats* usage_out,
                                    const GeminiHttpTimeouts* timeouts) {
    return GeminiGeneratePlainTextImpl(api_key_utf8, model_id_utf8, user_text_utf8, error_out,
                                       usage_out, timeouts);
}

std::string GeminiSupportedModelsCatalogJson() {
    boost::json::array models;
    auto add = [&models](const char* id, const char* name, const char* notes_utf8) {
        boost::json::object o;
        o["id"] = id;
        o["name"] = name;
        o["notes"] = std::string(notes_utf8);
        models.push_back(o);
    };
    add("gemini-3.1-pro-preview", "Gemini 3.1 Pro",
        u8"Линейка выше 2.5; превью, доступность и квоты — в AI Studio.");
    add("gemini-3.1-flash-lite-preview", "Gemini 3.1 Flash-Lite",
        u8"Облегчённая 3.1; превью.");
    add("gemini-3-flash-preview", "Gemini 3 Flash",
        u8"Основной Flash линейки 3 в API (имени gemini-3.0-flash нет). Не путать с несуществующим gemini-3.1-flash без суффикса.");
    add("gemini-2.5-pro", "Gemini 2.5 Pro", u8"Выше качество, ниже скорость; PDF и изображения.");
    add("gemini-2.5-flash", "Gemini 2.5 Flash",
        u8"PDF и изображения, быстрый ответ.");
    add("gemini-2.5-flash-lite", "Gemini 2.5 Flash-Lite",
        u8"Облегчённый вариант 2.5; быстрее/дешевле Flash.");

    boost::json::object root;
    root["defaultModelId"] = std::string(kGeminiDefaultModelId);
    root["namingNote"] = std::string(
        u8"Отдельных моделей gemini-3.0-flash и gemini-3.1-flash (без суффикса) в Gemini API нет. "
        u8"Flash для линейки 3 — gemini-3-flash-preview; облегчённый вариант 3.1 — gemini-3.1-flash-lite-preview. "
        u8"Для картинок в документации также фигурирует gemini-3.1-flash-image-preview.");
    root["propertyHint"] =
        std::string(u8"В свойство МодельGemini (GeminiModel) подставляйте значение поля id из массива models.");
    root["disclaimer"] = std::string(
        u8"Модели линейки Gemini 1.x и ниже 2.0 в список не включены (для generateContent не рекомендуются). "
        u8"Идентификаторы с суффиксом preview могут меняться. "
        u8"Фактический доступ — в https://aistudio.google.com/ и в документации Gemini API.");
    root["models"] = models;
    return boost::json::serialize(root);
}
