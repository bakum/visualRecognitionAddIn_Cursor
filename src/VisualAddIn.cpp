/*
 *  Modern Native AddIn
 *  Copyright (C) 2018  Infactum
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU Affero General Public License as
 *  published by the Free Software Foundation, either version 3 of the
 *  License, or (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Affero General Public License for more details.
 *
 *  You should have received a copy of the GNU Affero General Public License
 *  along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 */

#include "VisualAddIn.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <memory>
#include <optional>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include <boost/json.hpp>

#include "AnthropicMessagesClient.h"
#include "GeminiPdfClient.h"

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

constexpr int32_t kErrOk = 0;
constexpr int32_t kErrEmptyPdf = 1;
constexpr int32_t kErrInvalidBase64 = 4;
constexpr int32_t kErrWrongTypeBlob = 5;
constexpr int32_t kErrWrongTypeBase64 = 6;
constexpr int32_t kErrMissingApiKey = 7;
constexpr int32_t kErrAiFailure = 8;
constexpr int32_t kErrEmptyTextPrompt = 9;
constexpr int32_t kErrTextPromptTooLarge = 10;
constexpr int32_t kErrWrongTypeText = 11;
constexpr int32_t kErrInvalidEncryptedText = 12;
constexpr int32_t kErrUnknown = 99;
constexpr int32_t kGeminiFastProfileReceiveTimeoutMs = 20000;
constexpr int32_t kGeminiFastProfileTotalDeadlineMs = 30000;
constexpr size_t kGeminiRawPreviewMaxBytes = 4096U;

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

} // namespace

VisualAddIn::VisualAddIn()
    : last_error_code_storage_(std::make_shared<variant_t>(int32_t{0})),
      last_error_text_storage_(std::make_shared<variant_t>(std::string())),
      ai_studio_api_key_storage_(std::make_shared<variant_t>(std::string())),
      gemini_model_storage_(std::make_shared<variant_t>(std::string(kGeminiDefaultModelId))),
      gemini_receive_timeout_ms_storage_(std::make_shared<variant_t>(int32_t{45000})),
      gemini_total_deadline_ms_storage_(std::make_shared<variant_t>(int32_t{65000})),
      last_prompt_tokens_storage_(std::make_shared<variant_t>(0.0)),
      last_output_tokens_storage_(std::make_shared<variant_t>(0.0)),
      last_total_tokens_storage_(std::make_shared<variant_t>(0.0)),
      last_usage_json_storage_(std::make_shared<variant_t>(std::string("{}"))),
      last_gemini_raw_response_preview_storage_(std::make_shared<variant_t>(std::string())),
      ai_backend_storage_(std::make_shared<variant_t>(std::string("gemini"))),
      anthropic_api_key_storage_(std::make_shared<variant_t>(std::string())),
      anthropic_model_storage_(std::make_shared<variant_t>(std::string(kAnthropicDefaultModelId))),
      anthropic_max_output_tokens_storage_(std::make_shared<variant_t>(int32_t{8192})),
      last_anthropic_raw_response_preview_storage_(std::make_shared<variant_t>(std::string())) {
    AddProperty(
        u"Version",
        u"Версия",
        []() {
            static const auto storage =
                std::make_shared<variant_t>(std::string(VISUAL_ADDIN_VERSION));
            return storage;
        },
        nullptr);

    AddProperty(
        u"LastErrorCode",
        u"ПоследняяОшибка",
        [this]() {
            return last_error_code_storage_;
        },
        nullptr);
    AddProperty(
        u"LastErrorText",
        u"ТекстПоследнейОшибки",
        [this]() {
            return last_error_text_storage_;
        },
        nullptr);

    AddProperty(u"AiProvider", u"ПровайдерИИ", ai_backend_storage_);

    AddProperty(u"AIStudioApiKey", u"КлючAPIAIStudio", ai_studio_api_key_storage_);
    AddProperty(u"GeminiModel", u"МодельGemini", gemini_model_storage_);
    AddProperty(u"GeminiReceiveTimeoutMs", u"ТаймаутПолученияGeminiМс",
                gemini_receive_timeout_ms_storage_);
    AddProperty(u"GeminiTotalDeadlineMs", u"ОбщийДедлайнGeminiМс",
                gemini_total_deadline_ms_storage_);
    AddProperty(u"LastPromptTokens", u"ПоследниеВходящиеТокены", last_prompt_tokens_storage_);
    AddProperty(u"LastOutputTokens", u"ПоследниеИсходящиеТокены", last_output_tokens_storage_);
    AddProperty(u"LastTotalTokens", u"ПоследниеВсегоТокенов", last_total_tokens_storage_);
    AddProperty(u"LastGeminiUsageJson", u"ПоследняяСтатистикаТокеновJSON", last_usage_json_storage_);
    AddProperty(u"LastGeminiRawResponsePreview", u"ПоследнийСыройОтветGeminiПревью",
                last_gemini_raw_response_preview_storage_);
    AddProperty(u"AnthropicApiKey", u"КлючAPIAnthropic", anthropic_api_key_storage_);
    AddProperty(u"AnthropicModel", u"МодельAnthropic", anthropic_model_storage_);
    AddProperty(u"AnthropicMaxOutputTokens", u"МаксТокеновВыводаAnthropic",
                anthropic_max_output_tokens_storage_);
    AddProperty(u"LastAnthropicRawResponsePreview", u"ПоследнийСыройОтветAnthropicПревью",
                last_anthropic_raw_response_preview_storage_);

    // Перший аргумент AddMethod — англійська назва (lang=0), другий — російська (lang=1).
    // Два варіанти PDF викликають той самий шлях розбору (Gemini або Anthropic за ПровайдерИИ).
    AddMethod(u"ParsePrimaryDocumentPdf", u"РазобратьПервичныйДокументPdf", this,
              &VisualAddIn::ParsePrimaryDocumentPdfAi, {});
    AddMethod(u"EncodeToBase64", u"КодироватьВBase64", this, &VisualAddIn::EncodeToBase64, {});
    AddMethod(u"DecodeFromBase64", u"ДекодироватьИзBase64", this, &VisualAddIn::DecodeFromBase64,
              {});
    AddMethod(u"EncryptTextSimple", u"ШифроватьТекстПросто", this, &VisualAddIn::EncryptTextSimple,
              {});
    AddMethod(u"DecryptTextSimple", u"РасшифроватьТекстПросто", this,
              &VisualAddIn::DecryptTextSimple, {});
    AddMethod(u"EncodeKey", u"КодироватьКлюч", this, &VisualAddIn::EncodeKey, {});
    AddMethod(u"DecodeKey", u"ДекодироватьКлюч", this, &VisualAddIn::DecodeKey, {});
    AddMethod(u"ParsePrimaryDocumentPdfBase64", u"РазобратьПервичныйДокументPdfBase64", this,
              &VisualAddIn::ParsePrimaryDocumentPdfAiBase64, {});
    AddMethod(u"ParsePrimaryDocumentPdfGemini", u"РазобратьПервичныйДокументPdfИИ", this,
              &VisualAddIn::ParsePrimaryDocumentPdfAi, {});
    AddMethod(u"ParsePrimaryDocumentPdfGeminiBase64", u"РазобратьПервичныйДокументPdfИИBase64", this,
              &VisualAddIn::ParsePrimaryDocumentPdfAiBase64, {});
    AddMethod(u"ParsePrimaryDocumentImageGemini", u"РазобратьПервичныйДокументИзображениеИИ", this,
              &VisualAddIn::ParsePrimaryDocumentImageAi, {});
    AddMethod(u"ParsePrimaryDocumentImageGeminiBase64",
              u"РазобратьПервичныйДокументИзображениеИИBase64", this,
              &VisualAddIn::ParsePrimaryDocumentImageAiBase64, {});
    AddMethod(u"GenerateGeminiText", u"СгенерироватьТекстИИ", this, &VisualAddIn::GenerateGeminiText,
              {});
    AddMethod(u"GetSupportedGeminiModels", u"ПолучитьПоддерживаемыеМоделиGemini", this,
              &VisualAddIn::GetSupportedGeminiModels, {});
    AddMethod(u"GetSupportedAnthropicModels", u"ПолучитьПоддерживаемыеМоделиAnthropic", this,
              &VisualAddIn::GetSupportedAnthropicModels, {});
    AddMethod(u"UseFastGeminiTimeoutsProfile", u"ИспользоватьБыстрыйПрофильТаймаутовGemini", this,
              &VisualAddIn::UseFastGeminiTimeoutsProfile, {});
}

variant_t VisualAddIn::GetSupportedGeminiModels() {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    ResetUsageStats();
    return GeminiSupportedModelsCatalogJson();
}

variant_t VisualAddIn::GetSupportedAnthropicModels() {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    ResetUsageStats();
    return AnthropicSupportedModelsCatalogJson();
}

variant_t VisualAddIn::UseFastGeminiTimeoutsProfile() {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (gemini_receive_timeout_ms_storage_) {
        *gemini_receive_timeout_ms_storage_ =
            int32_t{kGeminiFastProfileReceiveTimeoutMs};
    }
    if (gemini_total_deadline_ms_storage_) {
        *gemini_total_deadline_ms_storage_ =
            int32_t{kGeminiFastProfileTotalDeadlineMs};
    }
    boost::json::object out;
    out["profile"] = "fast";
    out["receiveTimeoutMs"] = kGeminiFastProfileReceiveTimeoutMs;
    out["totalDeadlineMs"] = kGeminiFastProfileTotalDeadlineMs;
    return std::string(boost::json::serialize(out));
}

void VisualAddIn::ResetUsageStats() {
    if (last_prompt_tokens_storage_) {
        *last_prompt_tokens_storage_ = 0.0;
    }
    if (last_output_tokens_storage_) {
        *last_output_tokens_storage_ = 0.0;
    }
    if (last_total_tokens_storage_) {
        *last_total_tokens_storage_ = 0.0;
    }
    if (last_usage_json_storage_) {
        *last_usage_json_storage_ = std::string("{}");
    }
    if (last_gemini_raw_response_preview_storage_) {
        *last_gemini_raw_response_preview_storage_ = std::string();
    }
    if (last_anthropic_raw_response_preview_storage_) {
        *last_anthropic_raw_response_preview_storage_ = std::string();
    }
}

void VisualAddIn::SetUsageStats(const int64_t prompt_tokens,
                                const int64_t output_tokens,
                                const int64_t total_tokens,
                                const bool has_usage) {
    if (last_prompt_tokens_storage_) {
        *last_prompt_tokens_storage_ = has_usage ? static_cast<double>(prompt_tokens) : 0.0;
    }
    if (last_output_tokens_storage_) {
        *last_output_tokens_storage_ = has_usage ? static_cast<double>(output_tokens) : 0.0;
    }
    if (last_total_tokens_storage_) {
        *last_total_tokens_storage_ = has_usage ? static_cast<double>(total_tokens) : 0.0;
    }
    if (last_usage_json_storage_) {
        GeminiUsageStats usage;
        usage.prompt_tokens = prompt_tokens;
        usage.output_tokens = output_tokens;
        usage.total_tokens = total_tokens;
        usage.has_usage = has_usage;
        *last_usage_json_storage_ = UsageJsonFromStats(usage);
    }
}

variant_t VisualAddIn::ParsePrimaryDocumentAiFromBytes(const std::vector<char>& bytes,
                                                       const bool inline_as_pdf) {
    ResetUsageStats();
    if (bytes.empty()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrEmptyPdf,
                         inline_as_pdf ? std::string(u8"Пустые данные PDF.")
                                       : std::string(u8"Пустые данные изображения."));
        return JsonErrorObjectUtf8(inline_as_pdf ? "Empty PDF data" : "Empty image data");
    }

    GeminiHttpTimeouts timeouts;
    timeouts.receive_timeout_ms = GetIntPropertyOrDefault(
        gemini_receive_timeout_ms_storage_, 45000, 1000, 300000);
    timeouts.total_deadline_ms = GetIntPropertyOrDefault(
        gemini_total_deadline_ms_storage_, 65000, 5000, 600000);

    const bool use_anthropic = AiBackendIsAnthropic(ai_backend_storage_);
    std::string pipeline_err;
    GeminiUsageStats usage;
    std::string raw_response;
    std::string json_out;

    if (use_anthropic) {
        const std::string api_key = GetTrimmedUtf8Property(anthropic_api_key_storage_);
        if (api_key.empty()) {
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrMissingApiKey,
                             std::string(u8"Не задан ключ API Anthropic (КлючAPIAnthropic / AnthropicApiKey)."));
            variant_t err = JsonErrorObjectUtf8("Anthropic API key is empty");
            SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
            return err;
        }
        const std::string model_id = GetTrimmedUtf8Property(anthropic_model_storage_);
        if (LooksLikeGeminiModelId(model_id)) {
            variant_t err = JsonErrorObjectUtf8("Gemini model id used with Anthropic provider");
            SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
            return err;
        }
        const int max_out =
            GetIntPropertyOrDefault(anthropic_max_output_tokens_storage_, 32000, 4096, 128000);
        std::string_view mime = "application/pdf";
        if (!inline_as_pdf) {
            const std::optional<std::string> det = DetectRasterImageMimeForAi(bytes);
            if (!det.has_value()) {
                SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrAiFailure,
                                 std::string(u8"Формат изображения не поддерживается (ожидаются JPEG, PNG, "
                                             u8"GIF, WebP, BMP, TIFF)."));
                return JsonErrorObjectUtf8("Unsupported inline image format");
            }
            mime = *det;
        }
        json_out = AnthropicExtractPrimaryDocumentJson(api_key, model_id, mime, bytes, max_out,
                                                       pipeline_err, &usage, &timeouts, &raw_response);
        if (last_anthropic_raw_response_preview_storage_) {
            *last_anthropic_raw_response_preview_storage_ =
                MakeUtf8Preview(raw_response, kGeminiRawPreviewMaxBytes);
        }
    } else {
        const std::string api_key = GetTrimmedUtf8Property(ai_studio_api_key_storage_);
        if (api_key.empty()) {
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrMissingApiKey,
                             std::string(u8"Не задан ключ API (КлючAPIAIStudio / AIStudioApiKey)."));
            return JsonErrorObjectUtf8("Gemini API key is empty");
        }
        const std::string model_id = GetTrimmedUtf8Property(gemini_model_storage_);
        if (LooksLikeAnthropicModelId(model_id)) {
            variant_t err = JsonErrorObjectUtf8("Anthropic model id used with Gemini provider");
            SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
            return err;
        }
        json_out =
            inline_as_pdf
                ? GeminiExtractPrimaryDocumentJson(api_key, model_id, bytes, pipeline_err, &usage,
                                                     &timeouts, &raw_response)
                : GeminiExtractPrimaryDocumentJsonFromImageBytes(api_key, model_id, bytes, pipeline_err,
                                                                 &usage, &timeouts, &raw_response);
        if (last_gemini_raw_response_preview_storage_) {
            *last_gemini_raw_response_preview_storage_ =
                MakeUtf8Preview(raw_response, kGeminiRawPreviewMaxBytes);
        }
    }

    SetUsageStats(usage.prompt_tokens, usage.output_tokens, usage.total_tokens, usage.has_usage);
    if (!pipeline_err.empty()) {
        if (pipeline_err == "Invalid Gemini model id" || pipeline_err == "Invalid Anthropic model id"
            || pipeline_err == "PDF bytes passed to image method"
            || pipeline_err == "Unsupported inline image format" || pipeline_err == "Empty image data"
            || pipeline_err == "Invalid inline MIME type" || pipeline_err == "Empty model JSON body"
            || pipeline_err == "Anthropic request payload too large"
            || pipeline_err == "Anthropic output truncated (max_tokens)"
            || pipeline_err == "Gemini API key is empty" || pipeline_err == "Anthropic API key is empty") {
            const auto mapped = MapPipelineErrorToCodeAndText(pipeline_err);
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, mapped.first,
                             mapped.second);
        } else {
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrAiFailure,
                             (use_anthropic ? std::string(u8"Ошибка Anthropic API: ")
                                            : std::string(u8"Ошибка Gemini API: "))
                                 + pipeline_err);
        }
        return JsonErrorObjectUtf8(pipeline_err);
    }
    return json_out;
}

variant_t VisualAddIn::ParsePrimaryDocumentPdfAi(variant_t& pdf_blob) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::vector<char>>(pdf_blob)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBlob,
                         std::string(u8"Ожидались двоичные данные PDF (VTYPE_BLOB)."));
        throw std::runtime_error("Expected PDF as binary data (VTYPE_BLOB)");
    }
    variant_t r = ParsePrimaryDocumentAiFromBytes(std::get<std::vector<char>>(pdf_blob), true);
    SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, r);
    return r;
}

variant_t VisualAddIn::ParsePrimaryDocumentImageAi(variant_t& image_blob) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::vector<char>>(image_blob)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBlob,
                         std::string(u8"Ожидались двоичные данные изображения (VTYPE_BLOB)."));
        throw std::runtime_error("Expected image as binary data (VTYPE_BLOB)");
    }
    variant_t r = ParsePrimaryDocumentAiFromBytes(std::get<std::vector<char>>(image_blob), false);
    SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, r);
    return r;
}

variant_t VisualAddIn::EncodeToBase64(variant_t& bytes_blob) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::vector<char>>(bytes_blob)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBlob,
                         std::string(u8"Ожидались двоичные данные (VTYPE_BLOB)."));
        throw std::runtime_error("Expected binary data (VTYPE_BLOB)");
    }
    return EncodeBase64(std::get<std::vector<char>>(bytes_blob));
}

variant_t VisualAddIn::DecodeFromBase64(variant_t& base64_text) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(base64_text)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBase64,
                         std::string(u8"Ожидалась строка Base64 (VTYPE_PWSTR / UTF-8)."));
        throw std::runtime_error("Expected Base64 string (VTYPE_PWSTR / UTF-8 string)");
    }
    const std::optional<std::vector<char>> decoded = DecodeBase64(std::get<std::string>(base64_text));
    if (!decoded.has_value()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrInvalidBase64,
                         std::string(u8"Некорректная строка Base64."));
        variant_t err = JsonErrorObjectUtf8("Invalid Base64 string");
        SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
        return err;
    }
    return *decoded;
}

variant_t VisualAddIn::EncryptTextSimple(variant_t& plain_text) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(plain_text)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeText,
                         std::string(u8"Ожидалась строка UTF-8 (VTYPE_PWSTR)."));
        throw std::runtime_error("Expected text as UTF-8 string (VTYPE_PWSTR)");
    }
    return EncryptUtf8ToHex(std::get<std::string>(plain_text));
}

variant_t VisualAddIn::DecryptTextSimple(variant_t& encrypted_text) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(encrypted_text)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeText,
                         std::string(u8"Ожидалась шифрованная строка HEX (VTYPE_PWSTR)."));
        throw std::runtime_error("Expected HEX encrypted text as UTF-8 string (VTYPE_PWSTR)");
    }
    const std::optional<std::string> decrypted =
        DecryptHexToUtf8(std::get<std::string>(encrypted_text));
    if (!decrypted.has_value()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrInvalidEncryptedText,
                         std::string(u8"Некорректный формат шифрованного текста (ожидается HEX)."));
        return JsonErrorObjectUtf8("Invalid encrypted text format");
    }
    return *decrypted;
}

variant_t VisualAddIn::EncodeKey(variant_t& plain_text) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(plain_text)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeText,
                         std::string(u8"Ожидалась строка UTF-8 (VTYPE_PWSTR)."));
        throw std::runtime_error("Expected key as UTF-8 string (VTYPE_PWSTR)");
    }

    const std::string& input_text = std::get<std::string>(plain_text);
    variant_t bytes_blob = std::vector<char>(input_text.begin(), input_text.end());
    variant_t base64_text = EncodeToBase64(bytes_blob);
    return EncryptTextSimple(base64_text);
}

variant_t VisualAddIn::DecodeKey(variant_t& encoded_key) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(encoded_key)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeText,
                         std::string(u8"Ожидалась шифрованная строка UTF-8 (VTYPE_PWSTR)."));
        throw std::runtime_error("Expected encoded key as UTF-8 string (VTYPE_PWSTR)");
    }

    variant_t encrypted_text = std::get<std::string>(encoded_key);
    variant_t base64_text = DecryptTextSimple(encrypted_text);
    if (!std::holds_alternative<std::string>(base64_text)) {
        return base64_text;
    }

    variant_t decoded_blob = DecodeFromBase64(base64_text);
    if (!std::holds_alternative<std::vector<char>>(decoded_blob)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrInvalidEncryptedText,
                         std::string(u8"Не удалось декодировать Base64 после расшифровки ключа."));
        return JsonErrorObjectUtf8("Failed to decode Base64 from decrypted key");
    }

    const std::vector<char>& bytes = std::get<std::vector<char>>(decoded_blob);
    return std::string(bytes.begin(), bytes.end());
}

variant_t VisualAddIn::GenerateGeminiText(variant_t& prompt_utf8) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    ResetUsageStats();
    if (!std::holds_alternative<std::string>(prompt_utf8)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBase64,
                         std::string(u8"Ожидалась строка UTF-8 (VTYPE_PWSTR)."));
        throw std::runtime_error("Expected prompt as UTF-8 string (VTYPE_PWSTR)");
    }

    GeminiHttpTimeouts timeouts;
    timeouts.receive_timeout_ms = GetIntPropertyOrDefault(
        gemini_receive_timeout_ms_storage_, 45000, 1000, 300000);
    timeouts.total_deadline_ms = GetIntPropertyOrDefault(
        gemini_total_deadline_ms_storage_, 65000, 5000, 600000);

    const bool use_anthropic = AiBackendIsAnthropic(ai_backend_storage_);
    if (use_anthropic) {
        const std::string api_key = GetTrimmedUtf8Property(anthropic_api_key_storage_);
        if (api_key.empty()) {
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrMissingApiKey,
                             std::string(u8"Не задан ключ API Anthropic (КлючAPIAnthropic / AnthropicApiKey)."));
            variant_t err = JsonErrorObjectUtf8("Anthropic API key is empty");
            SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
            return err;
        }
        const std::string model_id = GetTrimmedUtf8Property(anthropic_model_storage_);
        if (LooksLikeGeminiModelId(model_id)) {
            variant_t err = JsonErrorObjectUtf8("Gemini model id used with Anthropic provider");
            SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
            return err;
        }
        const int max_out = GetIntPropertyOrDefault(anthropic_max_output_tokens_storage_, 8192, 1, 128000);
        std::string api_err;
        GeminiUsageStats usage;
        std::string raw_response;
        const std::string text_out =
            AnthropicGeneratePlainText(api_key, model_id, max_out, std::get<std::string>(prompt_utf8),
                                       api_err, &usage, &timeouts, &raw_response);
        if (last_anthropic_raw_response_preview_storage_) {
            *last_anthropic_raw_response_preview_storage_ =
                MakeUtf8Preview(raw_response, kGeminiRawPreviewMaxBytes);
        }
        SetUsageStats(usage.prompt_tokens, usage.output_tokens, usage.total_tokens, usage.has_usage);
        if (!api_err.empty()) {
            if (api_err == "Invalid Anthropic model id" || api_err == "Empty prompt text"
                || api_err == "Prompt text too large" || api_err == "Anthropic API key is empty") {
                const auto mapped = MapPipelineErrorToCodeAndText(api_err);
                SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, mapped.first,
                                 mapped.second);
            } else {
                SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrAiFailure,
                                 std::string(u8"Ошибка Anthropic API: ") + api_err);
            }
            return JsonErrorObjectUtf8(api_err);
        }
        return text_out;
    }

    const std::string api_key = GetTrimmedUtf8Property(ai_studio_api_key_storage_);
    if (api_key.empty()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrMissingApiKey,
                         std::string(u8"Не задан ключ API (КлючAPIAIStudio / AIStudioApiKey)."));
        variant_t err = JsonErrorObjectUtf8("Gemini API key is empty");
        SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
        return err;
    }
    const std::string model_id = GetTrimmedUtf8Property(gemini_model_storage_);
    if (LooksLikeAnthropicModelId(model_id)) {
        variant_t err = JsonErrorObjectUtf8("Anthropic model id used with Gemini provider");
        SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
        return err;
    }
    std::string gemini_err;
    GeminiUsageStats usage;
    std::string raw_response;
    const std::string text_out =
        GeminiGeneratePlainText(api_key, model_id, std::get<std::string>(prompt_utf8), gemini_err,
                                &usage, &timeouts, &raw_response);
    if (last_gemini_raw_response_preview_storage_) {
        *last_gemini_raw_response_preview_storage_ =
            MakeUtf8Preview(raw_response, kGeminiRawPreviewMaxBytes);
    }
    SetUsageStats(usage.prompt_tokens, usage.output_tokens, usage.total_tokens, usage.has_usage);
    if (!gemini_err.empty()) {
        if (gemini_err == "Invalid Gemini model id" || gemini_err == "Empty prompt text"
            || gemini_err == "Prompt text too large" || gemini_err == "Gemini API key is empty") {
            const auto mapped = MapPipelineErrorToCodeAndText(gemini_err);
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, mapped.first,
                             mapped.second);
        } else {
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrAiFailure,
                             std::string(u8"Ошибка Gemini API: ") + gemini_err);
        }
        return JsonErrorObjectUtf8(gemini_err);
    }
    return text_out;
}

variant_t VisualAddIn::ParsePrimaryDocumentPdfAiBase64(variant_t& pdf_base64) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(pdf_base64)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBase64,
                         std::string(u8"Ожидалась строка Base64 (VTYPE_PWSTR / UTF-8)."));
        throw std::runtime_error("Expected PDF as Base64 string (VTYPE_PWSTR / UTF-8 string)");
    }
    const std::string& b64 = std::get<std::string>(pdf_base64);
    const std::optional<std::vector<char>> pdf = DecodeBase64(b64);
    if (!pdf.has_value()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrInvalidBase64,
                         std::string(u8"Некорректная строка Base64."));
        variant_t err = JsonErrorObjectUtf8("Invalid Base64 string");
        SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
        return err;
    }
    variant_t blob = *pdf;
    return ParsePrimaryDocumentPdfAi(blob);
}

variant_t VisualAddIn::ParsePrimaryDocumentImageAiBase64(variant_t& image_base64) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::string>(image_base64)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBase64,
                         std::string(u8"Ожидалась строка Base64 (VTYPE_PWSTR / UTF-8)."));
        throw std::runtime_error("Expected image as Base64 string (VTYPE_PWSTR / UTF-8 string)");
    }
    const std::string& b64 = std::get<std::string>(image_base64);
    const std::optional<std::vector<char>> img = DecodeBase64(b64);
    if (!img.has_value()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrInvalidBase64,
                         std::string(u8"Некорректная строка Base64."));
        variant_t err = JsonErrorObjectUtf8("Invalid Base64 string");
        SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
        return err;
    }
    variant_t blob = *img;
    return ParsePrimaryDocumentImageAi(blob);
}

std::string VisualAddIn::extensionName() {
    return "VisualRecognition";
}
