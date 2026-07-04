#include "AddInUtils.h"

#include <algorithm>
#include <cctype>

#include <boost/json.hpp>

namespace addin {
namespace {

int Base64DecodeUnit(unsigned char c) {
    if (c >= 'A' && c <= 'Z') {
        return c - 'A';
    }
    if (c >= 'a' && c <= 'z') {
        return c - 'a' + 26;
    }
    if (c >= '0' && c <= '9') {
        return c - '0' + 52;
    }
    if (c == '+') {
        return 62;
    }
    if (c == '/') {
        return 63;
    }
    return -1;
}

constexpr char kSimpleCipherKey[] = "VisualRecognitionAddInSimpleKey";

char NibbleToHex(const unsigned value) {
    return static_cast<char>(value < 10U ? ('0' + value) : ('A' + (value - 10U)));
}

int HexToNibble(const char ch) {
    if (ch >= '0' && ch <= '9') {
        return ch - '0';
    }
    if (ch >= 'A' && ch <= 'F') {
        return ch - 'A' + 10;
    }
    if (ch >= 'a' && ch <= 'f') {
        return ch - 'a' + 10;
    }
    return -1;
}

unsigned char SimpleCipherByte(const unsigned char src, const size_t index) {
    const unsigned char key_byte =
        static_cast<unsigned char>(kSimpleCipherKey[index % (sizeof(kSimpleCipherKey) - 1U)]);
    const unsigned char salt = static_cast<unsigned char>((index * 31U) & 0xFFU);
    return static_cast<unsigned char>(src ^ key_byte ^ salt);
}

} // namespace

std::optional<std::vector<char>> DecodeBase64(std::string_view in) {
    std::vector<unsigned char> clean;
    clean.reserve(in.size());
    for (unsigned char ch : in) {
        if (ch == '=') {
            break;
        }
        if (std::isspace(static_cast<unsigned char>(ch)) != 0) {
            continue;
        }
        if (Base64DecodeUnit(ch) < 0) {
            return std::nullopt;
        }
        clean.push_back(ch);
    }
    const size_t n = clean.size();
    if (n % 4U == 1U) {
        return std::nullopt;
    }
    std::vector<char> out;
    out.reserve((n * 3U) / 4U);
    for (size_t i = 0; i < n; i += 4U) {
        const int a = Base64DecodeUnit(clean[i]);
        const int b = i + 1U < n ? Base64DecodeUnit(clean[i + 1U]) : -1;
        const int c = i + 2U < n ? Base64DecodeUnit(clean[i + 2U]) : -1;
        const int d = i + 3U < n ? Base64DecodeUnit(clean[i + 3U]) : -1;
        if (a < 0 || b < 0) {
            return std::nullopt;
        }
        if (i + 4U <= n) {
            if (c < 0 || d < 0) {
                return std::nullopt;
            }
            out.push_back(static_cast<char>((static_cast<unsigned>(a) << 2U)
                                            | (static_cast<unsigned>(b) >> 4U)));
            out.push_back(static_cast<char>(((static_cast<unsigned>(b) & 0xFU) << 4U)
                                            | (static_cast<unsigned>(c) >> 2U)));
            out.push_back(static_cast<char>(((static_cast<unsigned>(c) & 0x3U) << 6U)
                                            | static_cast<unsigned>(d)));
        } else if (i + 3U == n) {
            if (c < 0) {
                return std::nullopt;
            }
            out.push_back(static_cast<char>((static_cast<unsigned>(a) << 2U)
                                            | (static_cast<unsigned>(b) >> 4U)));
            out.push_back(static_cast<char>(((static_cast<unsigned>(b) & 0xFU) << 4U)
                                            | (static_cast<unsigned>(c) >> 2U)));
        } else if (i + 2U == n) {
            out.push_back(static_cast<char>((static_cast<unsigned>(a) << 2U)
                                            | (static_cast<unsigned>(b) >> 4U)));
        }
    }
    return out;
}

std::string EncodeBase64(const std::vector<char>& in) {
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

std::string EncryptUtf8ToHex(std::string_view plain_utf8) {
    std::string out;
    out.reserve(plain_utf8.size() * 2U);
    for (size_t i = 0; i < plain_utf8.size(); ++i) {
        const unsigned char encrypted =
            SimpleCipherByte(static_cast<unsigned char>(plain_utf8[i]), i);
        out.push_back(NibbleToHex((encrypted >> 4U) & 0xFU));
        out.push_back(NibbleToHex(encrypted & 0xFU));
    }
    return out;
}

std::optional<std::string> DecryptHexToUtf8(std::string_view encrypted_hex) {
    if ((encrypted_hex.size() % 2U) != 0U) {
        return std::nullopt;
    }
    std::string out;
    out.reserve(encrypted_hex.size() / 2U);
    for (size_t i = 0; i < encrypted_hex.size(); i += 2U) {
        const int hi = HexToNibble(encrypted_hex[i]);
        const int lo = HexToNibble(encrypted_hex[i + 1U]);
        if (hi < 0 || lo < 0) {
            return std::nullopt;
        }
        const unsigned char encrypted =
            static_cast<unsigned char>((static_cast<unsigned>(hi) << 4U) | static_cast<unsigned>(lo));
        out.push_back(static_cast<char>(SimpleCipherByte(encrypted, i / 2U)));
    }
    return out;
}

std::optional<std::string> ExtractErrorFromResultJson(const std::string& json_utf8) {
    try {
        const boost::json::value v = boost::json::parse(json_utf8);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& o = v.as_object();
        const auto it = o.find("error");
        if (it == o.end() || !it->value().is_string()) {
            return std::nullopt;
        }
        return std::string(it->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}

void SetLastErrorPair(const std::shared_ptr<variant_t>& code_storage,
                      const std::shared_ptr<variant_t>& text_storage,
                      const int32_t code,
                      std::string text_utf8) {
    if (code_storage) {
        *code_storage = code;
    }
    if (text_storage) {
        *text_storage = std::move(text_utf8);
    }
}

void ClearLastErrorPair(const std::shared_ptr<variant_t>& code_storage,
                        const std::shared_ptr<variant_t>& text_storage) {
    SetLastErrorPair(code_storage, text_storage, kErrOk, std::string());
}

std::pair<int32_t, std::string> MapPipelineErrorToCodeAndText(const std::string& error_field_utf8) {
    if (error_field_utf8 == "Empty PDF data") {
        return {kErrEmptyPdf, std::string(u8"Пустые данные PDF.")};
    }
    if (error_field_utf8 == "Invalid Base64 string") {
        return {kErrInvalidBase64, std::string(u8"Некорректная строка Base64.")};
    }
    if (error_field_utf8 == "Gemini API key is empty" || error_field_utf8 == "API key is empty") {
        return {kErrMissingApiKey,
                std::string(u8"Не задан ключ API (свойства КлючAPIAIStudio / AIStudioApiKey).")};
    }
    if (error_field_utf8 == "Anthropic API key is empty") {
        return {kErrMissingApiKey,
                std::string(u8"Не задан ключ API Anthropic (КлючAPIAnthropic / AnthropicApiKey).")};
    }
    if (error_field_utf8 == "Invalid Gemini model id") {
        return {kErrAiFailure,
                std::string(u8"Некорректный идентификатор модели Gemini (свойство МодельGemini / GeminiModel): "
                            u8"допустимы латиница, цифры, «.», «-», «_».")};
    }
    if (error_field_utf8 == "Invalid Anthropic model id") {
        return {kErrAiFailure,
                std::string(u8"Некорректный идентификатор модели Anthropic (свойство МодельAnthropic / "
                            u8"AnthropicModel): допустимы латиница, цифры, «.», «-», «_».")};
    }
    if (error_field_utf8 == "Empty image data") {
        return {kErrEmptyPdf, std::string(u8"Пустые данные изображения.")};
    }
    if (error_field_utf8 == "PDF bytes passed to image method") {
        return {kErrAiFailure,
                std::string(u8"Передан PDF; для PDF используйте РазобратьПервичныйДокументPdf (или …PdfИИ).")};
    }
    if (error_field_utf8 == "Unsupported inline image format") {
        return {kErrAiFailure,
                std::string(u8"Формат изображения не поддерживается (ожидаются JPEG, PNG, GIF, WebP, BMP, TIFF).")};
    }
    if (error_field_utf8 == "Empty prompt text") {
        return {kErrEmptyTextPrompt, std::string(u8"Пустой текст запроса (после обрезки пробелов).")};
    }
    if (error_field_utf8 == "Prompt text too large") {
        return {kErrTextPromptTooLarge,
                std::string(u8"Текст запроса слишком длинный (лимит 2 МиБ UTF-8).")};
    }
    if (error_field_utf8 == "Empty model JSON body") {
        return {kErrAiFailure, std::string(u8"Модель вернула пустое тело JSON.")};
    }
    if (error_field_utf8 == "Anthropic request payload too large") {
        return {kErrAiFailure, std::string(u8"Запрос к Anthropic слишком большой (лимит ~32 МиБ).")};
    }
    if (error_field_utf8 == "Anthropic output truncated (max_tokens)") {
        return {kErrAiFailure,
                std::string(u8"Ответ Anthropic обрезан по max_tokens; увеличьте МаксТокеновВыводаAnthropic.")};
    }
    if (error_field_utf8 == "Invalid inline MIME type") {
        return {kErrAiFailure, std::string(u8"Недопустимый MIME-тип вложения для ИИ-запроса.")};
    }
    if (error_field_utf8 == "Anthropic model id used with Gemini provider") {
        return {kErrAiFailure,
                std::string(u8"Указана модель Anthropic (Claude) в свойстве МодельGemini при провайдере "
                            u8"Gemini. Задайте ПровайдерИИ / AiProvider = anthropic (или антропик) и "
                            u8"перенесите идентификатор в МодельAnthropic / AnthropicModel.")};
    }
    if (error_field_utf8 == "Gemini model id used with Anthropic provider") {
        return {kErrAiFailure,
                std::string(u8"Указана модель Gemini в свойстве МодельAnthropic при провайдере Anthropic. "
                            u8"Используйте идентификатор Claude или верните ПровайдерИИ = gemini.")};
    }
    return {kErrUnknown, error_field_utf8};
}

void SyncLastErrorFromParseResult(const std::shared_ptr<variant_t>& code_storage,
                                  const std::shared_ptr<variant_t>& text_storage,
                                  const variant_t& r) {
    if (!code_storage || !text_storage) {
        return;
    }
    if (!std::holds_alternative<std::string>(r)) {
        return;
    }
    const auto& json = std::get<std::string>(r);
    const std::optional<std::string> err = ExtractErrorFromResultJson(json);
    if (!err.has_value()) {
        return;
    }
    const auto mapped = MapPipelineErrorToCodeAndText(*err);
    SetLastErrorPair(code_storage, text_storage, mapped.first, std::string(mapped.second));
}

std::string GetTrimmedUtf8Property(const std::shared_ptr<variant_t>& storage) {
    if (!storage || !std::holds_alternative<std::string>(*storage)) {
        return {};
    }
    std::string k = std::get<std::string>(*storage);
    while (!k.empty() && (static_cast<unsigned char>(k.front()) <= ' ')) {
        k.erase(k.begin());
    }
    while (!k.empty() && (static_cast<unsigned char>(k.back()) <= ' ')) {
        k.pop_back();
    }
    return k;
}

std::string AlphanumericAsciiLowerKey(std::string_view s) {
    std::string key;
    for (unsigned char ch : s) {
        if (std::isalnum(ch) != 0) {
            key.push_back(static_cast<char>(std::tolower(ch)));
        }
    }
    return key;
}

bool LooksLikeAnthropicModelId(const std::string& model_id) {
    std::string m = model_id;
    while (!m.empty() && (static_cast<unsigned char>(m.front()) <= ' ')) {
        m.erase(m.begin());
    }
    while (!m.empty() && (static_cast<unsigned char>(m.back()) <= ' ')) {
        m.pop_back();
    }
    for (size_t i = 0; i < m.size(); ++i) {
        unsigned char uc = static_cast<unsigned char>(m[i]);
        if (uc >= static_cast<unsigned char>('A') && uc <= static_cast<unsigned char>('Z')) {
            m[i] = static_cast<char>(uc - static_cast<unsigned char>('A') + static_cast<unsigned char>('a'));
        }
    }
    if (m.rfind("claude-", 0) == 0) {
        return true;
    }
    const std::string k = AlphanumericAsciiLowerKey(m);
    return k == "haiku45" || k == "sonnet45" || k == "sonnet35" || k == "haiku35" || k == "opus3"
        || k == "claudehaiku45" || k == "claudesonnet45" || k == "claude35sonnet" || k == "claude35haiku"
        || k == "claude3opus" || k == "claude45sonnet";
}

bool LooksLikeGeminiModelId(const std::string& model_id) {
    std::string m = model_id;
    while (!m.empty() && (static_cast<unsigned char>(m.front()) <= ' ')) {
        m.erase(m.begin());
    }
    while (!m.empty() && (static_cast<unsigned char>(m.back()) <= ' ')) {
        m.pop_back();
    }
    for (size_t i = 0; i < m.size(); ++i) {
        unsigned char uc = static_cast<unsigned char>(m[i]);
        if (uc >= static_cast<unsigned char>('A') && uc <= static_cast<unsigned char>('Z')) {
            m[i] = static_cast<char>(uc - static_cast<unsigned char>('A') + static_cast<unsigned char>('a'));
        }
    }
    if (m.rfind("gemini-", 0) == 0) {
        return true;
    }
    const std::string k = AlphanumericAsciiLowerKey(m);
    return k == "gemini3pro" || k == "gemini31pro" || k == "gemini3flashlite" || k == "gemini31flashlite"
        || k == "gemini3flash" || k == "gemini31flash";
}

void StripUtf8BomAndNbspEdgesInPlace(std::string& v) {
    while (v.size() >= 3U && static_cast<unsigned char>(v[0]) == 0xEFU
           && static_cast<unsigned char>(v[1]) == 0xBBU && static_cast<unsigned char>(v[2]) == 0xBFU) {
        v.erase(0, 3);
    }
    while (v.size() >= 2U && static_cast<unsigned char>(v[0]) == 0xC2U
           && static_cast<unsigned char>(v[1]) == 0xA0U) {
        v.erase(0, 2);
    }
    while (v.size() >= 2U && static_cast<unsigned char>(v[v.size() - 2U]) == 0xC2U
           && static_cast<unsigned char>(v[v.size() - 1U]) == 0xA0U) {
        v.resize(v.size() - 2U);
    }
}

bool AiBackendIsAnthropic(const std::shared_ptr<variant_t>& ai_backend_storage) {
    if (!ai_backend_storage || !std::holds_alternative<std::string>(*ai_backend_storage)) {
        return false;
    }
    std::string v = std::get<std::string>(*ai_backend_storage);
    while (!v.empty() && (static_cast<unsigned char>(v.front()) <= ' ')) {
        v.erase(v.begin());
    }
    while (!v.empty() && (static_cast<unsigned char>(v.back()) <= ' ')) {
        v.pop_back();
    }
    StripUtf8BomAndNbspEdgesInPlace(v);
    if (v == std::string(u8"антропик") || v == std::string(u8"Антропик") || v == std::string(u8"АНТРОПИК")
        || v == std::string(u8"клод") || v == std::string(u8"Клод") || v == std::string(u8"КЛОД")) {
        return true;
    }
    const std::string k = AlphanumericAsciiLowerKey(v);
    constexpr std::string_view kAnth = "anthropic";
    constexpr std::string_view kClaude = "claude";
    if (k.size() >= kAnth.size() && k.compare(0, kAnth.size(), kAnth) == 0) {
        return true;
    }
    if (k.size() >= kClaude.size() && k.compare(0, kClaude.size(), kClaude) == 0) {
        return true;
    }
    return false;
}

int GetIntPropertyOrDefault(const std::shared_ptr<variant_t>& storage,
                            const int default_value,
                            const int min_value,
                            const int max_value) {
    if (!storage) {
        return default_value;
    }
    int value = default_value;
    if (std::holds_alternative<int32_t>(*storage)) {
        value = static_cast<int>(std::get<int32_t>(*storage));
    } else if (std::holds_alternative<double>(*storage)) {
        value = static_cast<int>(std::get<double>(*storage));
    } else {
        return default_value;
    }
    if (value < min_value) {
        return min_value;
    }
    if (value > max_value) {
        return max_value;
    }
    return value;
}

variant_t JsonErrorObjectUtf8(const std::string& error_utf8) {
    boost::json::object o;
    o["error"] = error_utf8;
    return std::string(boost::json::serialize(o));
}

std::string UsageJsonFromStats(const GeminiUsageStats& usage) {
    if (!usage.has_usage) {
        return "{}";
    }
    boost::json::object o;
    o["promptTokenCount"] = usage.prompt_tokens;
    o["candidatesTokenCount"] = usage.output_tokens;
    o["totalTokenCount"] = usage.total_tokens;
    return std::string(boost::json::serialize(o));
}

std::string MakeUtf8Preview(const std::string& s, const size_t max_bytes) {
    if (s.size() <= max_bytes) {
        return s;
    }
    return s.substr(0, max_bytes) + "...(truncated)";
}

std::optional<std::string> DetectRasterImageMimeForAi(const std::vector<char>& b) {
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

} // namespace addin
