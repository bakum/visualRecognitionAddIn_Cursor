#include "GeminiPdfClient.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>
#include <winhttp.h>

#include <boost/json.hpp>

#include <algorithm>
#include <cctype>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

namespace {

constexpr wchar_t kHost[] = L"generativelanguage.googleapis.com";
constexpr wchar_t kPathPrefix[] = L"/v1beta/models/";
constexpr wchar_t kPathSuffix[] = L":generateContent";

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
    const int n = MultiByteToWideChar(CP_UTF8, 0, u8.data(), static_cast<int>(u8.size()), nullptr, 0);
    if (n <= 0) {
        return {};
    }
    std::wstring w(static_cast<size_t>(n), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, u8.data(), static_cast<int>(u8.size()), w.data(), n);
    return w;
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
        props["quantity"] = boost::json::object({{"type", "string"}});
        props["unit"] = boost::json::object({{"type", "string"}});
        props["price"] = boost::json::object({{"type", "string"}});
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

    boost::json::object root_props;
    root_props["counterparty"] = counterparty;
    root_props["contract"] = contract;
    root_props["invoiceNumber"] = boost::json::object({{"type", "string"}});
    root_props["bankDetails"] = bank;
    root_props["lineItems"] = line_items;
    root_props["rawText"] = boost::json::object({{"type", "string"}});

    boost::json::object schema;
    schema["type"] = "object";
    schema["properties"] = root_props;
    boost::json::array req;
    req.push_back("counterparty");
    req.push_back("contract");
    req.push_back("invoiceNumber");
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
        "the picture. The content may be a primary accounting document (invoice, bill, act, waybill, "
        "receipt, etc.) or another business graphic document; visible text may be Ukrainian, Russian, "
        "English, or other languages shown. Extract all data relevant to the schema and return a single "
        "JSON object that strictly follows the provided JSON schema. Use empty string \"\" for unknown "
        "fields. Preserve the original language and spelling from the document for text values. For "
        "lineItems, include every product/service row with name, sku, quantity, unit, price, amount when "
        "visible. Field sku is the article / vendor code / SKU when the document shows it (labels like "
        "Артикул, Артикул постачальника, Код товару, SKU, Article, Part number, Cat. no., etc.); use \"\" "
        "if no separate article column or value exists. Omit other fields only if absent. In counterparty: "
        "supplierInn is the tax id of the "
        "supplier/seller/issuer (e.g. ЄДРПОУ, ДРФО, ИНН next to Постачальник/Продавець/Виконавець); "
        "buyerInn is the buyer's tax id when shown (Покупець/Замовник). supplierKpp and buyerKpp are "
        "Russian КПП when present; otherwise \"\". supplierOkpo and buyerOkpo: in Russian "
        "documents use label ОКПО (8 digits for organizations, 10 for sole proprietors); in Ukrainian "
        "documents the same role is ЄДРПОУ (typically 8 digits for legal entities) — put that value "
        "in supplierOkpo/buyerOkpo when it is the party's registration/statistical code, even if also "
        "in supplierInn. If absent, \"\". Do not put the buyer's code into supplierOkpo.";

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
    "asks for plain or compact output.";

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
                    unsigned long& status_out,
                    std::string& response_body_out,
                    std::string& error_out) {
    const std::wstring path =
        std::wstring(kPathPrefix) + Utf8ToWide(model_id_utf8) + kPathSuffix;
    const std::wstring wkey = Utf8ToWide(api_key_utf8);

    HINTERNET session =
        WinHttpOpen(L"VisualRecognitionAddIn/1.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
                    WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
    if (!session) {
        error_out = "Gemini API network error (WinHttpOpen)";
        return false;
    }

    bool ok = false;
    HINTERNET connect =
        WinHttpConnect(session, kHost, INTERNET_DEFAULT_HTTPS_PORT, 0);
    if (!connect) {
        error_out = "Gemini API network error (WinHttpConnect)";
        WinHttpCloseHandle(session);
        return false;
    }

    HINTERNET request =
        WinHttpOpenRequest(connect, L"POST", path.c_str(), nullptr, WINHTTP_NO_REFERER,
                           WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
    if (!request) {
        error_out = "Gemini API network error (WinHttpOpenRequest)";
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        return false;
    }

    const std::wstring hdr_key = L"x-goog-api-key: " + wkey;
    std::wstring headers = hdr_key + L"\r\nContent-Type: application/json\r\n";

    const BOOL sent =
        WinHttpSendRequest(request, headers.c_str(), static_cast<DWORD>(-1),
                           const_cast<char*>(body_utf8.data()), static_cast<DWORD>(body_utf8.size()),
                           static_cast<DWORD>(body_utf8.size()), 0);

    if (!sent) {
        error_out = "Gemini API network error (WinHttpSendRequest)";
        WinHttpCloseHandle(request);
        WinHttpCloseHandle(connect);
        WinHttpCloseHandle(session);
        return false;
    }

    if (!WinHttpReceiveResponse(request, nullptr)) {
        error_out = "Gemini API network error (WinHttpReceiveResponse)";
        WinHttpCloseHandle(request);
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
    for (;;) {
        DWORD avail = 0;
        if (!WinHttpQueryDataAvailable(request, &avail)) {
            error_out = "Gemini API network error (WinHttpQueryDataAvailable)";
            WinHttpCloseHandle(request);
            WinHttpCloseHandle(connect);
            WinHttpCloseHandle(session);
            return false;
        }
        if (avail == 0) {
            break;
        }
        std::vector<char> chunk(static_cast<size_t>(avail));
        DWORD read = 0;
        if (!WinHttpReadData(request, chunk.data(), avail, &read)) {
            error_out = "Gemini API network error (WinHttpReadData)";
            WinHttpCloseHandle(request);
            WinHttpCloseHandle(connect);
            WinHttpCloseHandle(session);
            return false;
        }
        response_body_out.append(chunk.data(), read);
    }

    ok = true;
    WinHttpCloseHandle(request);
    WinHttpCloseHandle(connect);
    WinHttpCloseHandle(session);
    return ok;
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

std::string GeminiExtractPrimaryDocumentJsonImpl(const std::string& api_key_utf8,
                                                 const std::string& model_id_utf8,
                                                 std::string_view mime_type,
                                                 const std::vector<char>& bytes,
                                                 std::string& error_out) {
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
    if (!HttpPostGemini(api_key_utf8, model, body, http_status, raw_response, error_out)) {
        return {};
    }

    if (http_status != 200UL) {
        const auto api_err = ExtractGeminiApiErrorMessage(raw_response);
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Gemini API HTTP " + std::to_string(http_status);
        }
        return {};
    }

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
        const boost::json::value parsed = boost::json::parse(inner);
        (void)parsed;
    } catch (const std::exception& e) {
        error_out = std::string("Gemini JSON parse error: ") + e.what();
        return {};
    }

    return inner;
}

std::string GeminiGeneratePlainTextImpl(const std::string& api_key_utf8,
                                          const std::string& model_id_utf8,
                                          const std::string& user_text_utf8,
                                          std::string& error_out) {
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
    if (!HttpPostGemini(api_key_utf8, model, body, http_status, raw_response, error_out)) {
        return {};
    }

    if (http_status != 200UL) {
        const auto api_err = ExtractGeminiApiErrorMessage(raw_response);
        if (api_err.has_value()) {
            error_out = *api_err;
        } else {
            error_out = "Gemini API HTTP " + std::to_string(http_status);
        }
        return {};
    }

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
                                             std::string& error_out) {
    return GeminiExtractPrimaryDocumentJsonImpl(api_key_utf8, model_id_utf8, "application/pdf",
                                                pdf_bytes, error_out);
}

std::string GeminiExtractPrimaryDocumentJsonFromImageBytes(const std::string& api_key_utf8,
                                                           const std::string& model_id_utf8,
                                                           const std::vector<char>& image_bytes,
                                                           std::string& error_out) {
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
                                                error_out);
}

std::string GeminiGeneratePlainText(const std::string& api_key_utf8,
                                    const std::string& model_id_utf8,
                                    const std::string& user_text_utf8,
                                    std::string& error_out) {
    return GeminiGeneratePlainTextImpl(api_key_utf8, model_id_utf8, user_text_utf8, error_out);
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
    add("gemini-3.1-pro-preview", "Gemini 3.1 Pro (preview)",
        u8"Линейка выше 2.5; превью, доступность и квоты — в AI Studio.");
    add("gemini-3.1-flash-lite-preview", "Gemini 3.1 Flash-Lite (preview)",
        u8"Облегчённая 3.1; превью. Идентификатор по умолчанию в компоненте.");
    add("gemini-3-flash-preview", "Gemini 3 Flash (preview)",
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
