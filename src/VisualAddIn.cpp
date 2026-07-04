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

#include "AddInUtils.h"

using namespace addin;


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
        // Хранить MIME в std::string, а не string_view на *det: optional уничтожается после if,
        // иначе AnthropicExtractPrimaryDocumentJson получает висячий string_view → ложная ошибка MIME.
        std::string inline_mime{"application/pdf"};
        if (!inline_as_pdf) {
            const std::optional<std::string> det = DetectRasterImageMimeForAi(bytes);
            if (!det.has_value()) {
                SetLastErrorPair(last_error_code_storage_, last_error_text_storage_, kErrAiFailure,
                                 std::string(u8"Формат изображения не поддерживается (ожидаются JPEG, PNG, "
                                             u8"GIF, WebP, BMP, TIFF)."));
                return JsonErrorObjectUtf8("Unsupported inline image format");
            }
            inline_mime = *det;
        }
        json_out = AnthropicExtractPrimaryDocumentJson(api_key, model_id, inline_mime, bytes, max_out,
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
