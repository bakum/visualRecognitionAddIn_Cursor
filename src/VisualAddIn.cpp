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

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#include <cctype>
#include <cstdint>
#include <memory>
#include <optional>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

#include <boost/json.hpp>

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
constexpr int32_t kErrUnknown = 99;

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
    if (error_field_utf8 == "API key is empty") {
        return {kErrMissingApiKey,
                std::string(u8"Не задан ключ API (свойства КлючAPIAIStudio / AIStudioApiKey).")};
    }
    if (error_field_utf8 == "Invalid Gemini model id") {
        return {kErrAiFailure,
                std::string(u8"Некорректный идентификатор модели Gemini (свойство МодельGemini / GeminiModel): "
                            u8"допустимы латиница, цифры, «.», «-», «_».")};
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

} // namespace

VisualAddIn::VisualAddIn()
    : last_error_code_storage_(std::make_shared<variant_t>(int32_t{0})),
      last_error_text_storage_(std::make_shared<variant_t>(std::string())),
      ai_studio_api_key_storage_(std::make_shared<variant_t>(std::string())),
      gemini_model_storage_(std::make_shared<variant_t>(std::string(kGeminiDefaultModelId))),
      last_prompt_tokens_storage_(std::make_shared<variant_t>(0.0)),
      last_output_tokens_storage_(std::make_shared<variant_t>(0.0)),
      last_total_tokens_storage_(std::make_shared<variant_t>(0.0)),
      last_usage_json_storage_(std::make_shared<variant_t>(std::string("{}"))) {
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

    AddProperty(u"AIStudioApiKey", u"КлючAPIAIStudio", ai_studio_api_key_storage_);
    AddProperty(u"GeminiModel", u"МодельGemini", gemini_model_storage_);
    AddProperty(u"LastPromptTokens", u"ПоследниеВходящиеТокены", last_prompt_tokens_storage_);
    AddProperty(u"LastOutputTokens", u"ПоследниеИсходящиеТокены", last_output_tokens_storage_);
    AddProperty(u"LastTotalTokens", u"ПоследниеВсегоТокенов", last_total_tokens_storage_);
    AddProperty(u"LastGeminiUsageJson", u"ПоследняяСтатистикаТокеновJSON", last_usage_json_storage_);

    // Перший аргумент AddMethod — англійська назва (lang=0), другий — російська (lang=1).
    // Два варіанти PDF викликають той самий Gemini-шлях (старий код 1С без «ИИ» у імені).
    AddMethod(u"ParsePrimaryDocumentPdf", u"РазобратьПервичныйДокументPdf", this,
              &VisualAddIn::ParsePrimaryDocumentPdfAi, {});
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
}

variant_t VisualAddIn::GetSupportedGeminiModels() {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    ResetUsageStats();
    return GeminiSupportedModelsCatalogJson();
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

variant_t VisualAddIn::ParsePrimaryDocumentGeminiFromBytes(const std::vector<char>& bytes,
                                                           const bool inline_as_pdf) {
    ResetUsageStats();
    if (bytes.empty()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrEmptyPdf,
                         inline_as_pdf ? std::string(u8"Пустые данные PDF.")
                                       : std::string(u8"Пустые данные изображения."));
        return JsonErrorObjectUtf8(inline_as_pdf ? "Empty PDF data" : "Empty image data");
    }
    const std::string api_key = GetTrimmedUtf8Property(ai_studio_api_key_storage_);
    if (api_key.empty()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrMissingApiKey,
                         std::string(u8"Не задан ключ API (КлючAPIAIStudio / AIStudioApiKey)."));
        return JsonErrorObjectUtf8("API key is empty");
    }
    const std::string model_id = GetTrimmedUtf8Property(gemini_model_storage_);
    std::string gemini_err;
    GeminiUsageStats usage;
    const std::string json =
        inline_as_pdf
            ? GeminiExtractPrimaryDocumentJson(api_key, model_id, bytes, gemini_err, &usage)
            : GeminiExtractPrimaryDocumentJsonFromImageBytes(api_key, model_id, bytes, gemini_err,
                                                             &usage);
    SetUsageStats(usage.prompt_tokens, usage.output_tokens, usage.total_tokens, usage.has_usage);
    if (!gemini_err.empty()) {
        if (gemini_err == "Invalid Gemini model id") {
            const auto mapped = MapPipelineErrorToCodeAndText(gemini_err);
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, mapped.first,
                             mapped.second);
        } else if (gemini_err == "PDF bytes passed to image method"
                   || gemini_err == "Unsupported inline image format"
                   || gemini_err == "Empty image data") {
            const auto mapped = MapPipelineErrorToCodeAndText(gemini_err);
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, mapped.first,
                             mapped.second);
        } else {
            SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrAiFailure,
                             std::string(u8"Ошибка Gemini API: ") + gemini_err);
        }
        return JsonErrorObjectUtf8(gemini_err);
    }
    return json;
}

variant_t VisualAddIn::ParsePrimaryDocumentPdfAi(variant_t& pdf_blob) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    if (!std::holds_alternative<std::vector<char>>(pdf_blob)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBlob,
                         std::string(u8"Ожидались двоичные данные PDF (VTYPE_BLOB)."));
        throw std::runtime_error("Expected PDF as binary data (VTYPE_BLOB)");
    }
    variant_t r = ParsePrimaryDocumentGeminiFromBytes(std::get<std::vector<char>>(pdf_blob), true);
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
    variant_t r = ParsePrimaryDocumentGeminiFromBytes(std::get<std::vector<char>>(image_blob), false);
    SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, r);
    return r;
}

variant_t VisualAddIn::GenerateGeminiText(variant_t& prompt_utf8) {
    ClearLastErrorPair(last_error_code_storage_, last_error_text_storage_);
    ResetUsageStats();
    if (!std::holds_alternative<std::string>(prompt_utf8)) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrWrongTypeBase64,
                         std::string(u8"Ожидалась строка UTF-8 (VTYPE_PWSTR)."));
        throw std::runtime_error("Expected prompt as UTF-8 string (VTYPE_PWSTR)");
    }
    const std::string api_key = GetTrimmedUtf8Property(ai_studio_api_key_storage_);
    if (api_key.empty()) {
        SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrMissingApiKey,
                         std::string(u8"Не задан ключ API (КлючAPIAIStudio / AIStudioApiKey)."));
        variant_t err = JsonErrorObjectUtf8("API key is empty");
        SyncLastErrorFromParseResult(last_error_code_storage_, last_error_text_storage_, err);
        return err;
    }
    const std::string model_id = GetTrimmedUtf8Property(gemini_model_storage_);
    std::string gemini_err;
    GeminiUsageStats usage;
    const std::string text_out =
        GeminiGeneratePlainText(api_key, model_id, std::get<std::string>(prompt_utf8), gemini_err,
                                &usage);
    SetUsageStats(usage.prompt_tokens, usage.output_tokens, usage.total_tokens, usage.has_usage);
    if (!gemini_err.empty()) {
        if (gemini_err == "Invalid Gemini model id" || gemini_err == "Empty prompt text"
            || gemini_err == "Prompt text too large") {
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
