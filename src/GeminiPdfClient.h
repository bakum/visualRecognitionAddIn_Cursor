#ifndef GEMINI_PDF_CLIENT_H
#define GEMINI_PDF_CLIENT_H

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

/// Идентификатор модели по умолчанию (если из 1С передана пустая строка).
inline constexpr char kGeminiDefaultModelId[] = "gemini-3-flash-preview";

struct GeminiUsageStats {
    int64_t prompt_tokens = 0;
    int64_t output_tokens = 0;
    int64_t total_tokens = 0;
    bool has_usage = false;
};

struct GeminiHttpTimeouts {
    int receive_timeout_ms = 45000;
    int total_deadline_ms = 65000;
};

/// PDF (application/pdf) → Gemini; структура как у локального разбора первички.
/// model_id_utf8 — имя модели API; пусто после trim — kGeminiDefaultModelId.
std::string GeminiExtractPrimaryDocumentJson(const std::string& api_key_utf8,
                                             const std::string& model_id_utf8,
                                             const std::vector<char>& pdf_bytes,
                                             std::string& error_out,
                                             GeminiUsageStats* usage_out = nullptr,
                                             const GeminiHttpTimeouts* timeouts = nullptr,
                                             std::string* raw_response_out = nullptr);

/// Растровое изображение (JPEG/PNG/GIF/WebP/BMP/TIFF по сигнатуре). PDF не принимается — для PDF используйте PdfИИ.
std::string GeminiExtractPrimaryDocumentJsonFromImageBytes(const std::string& api_key_utf8,
                                                           const std::string& model_id_utf8,
                                                           const std::vector<char>& image_bytes,
                                                           std::string& error_out,
                                                           GeminiUsageStats* usage_out = nullptr,
                                                           const GeminiHttpTimeouts* timeouts = nullptr,
                                                           std::string* raw_response_out = nullptr);

/// Произвольный текстовый запрос к Gemini без схемы JSON и без разбора первички; ответ — обычная строка модели (UTF-8).
std::string GeminiGeneratePlainText(const std::string& api_key_utf8,
                                    const std::string& model_id_utf8,
                                    const std::string& user_text_utf8,
                                    std::string& error_out,
                                    GeminiUsageStats* usage_out = nullptr,
                                    const GeminiHttpTimeouts* timeouts = nullptr,
                                    std::string* raw_response_out = nullptr);

/// Каталог рекомендуемых моделей (UTF-8 JSON): defaultModelId, models[{id,name,notes}], подсказки. Ключ API не нужен.
std::string GeminiSupportedModelsCatalogJson();

/// Текст пользовательского сообщения (инструкция + JSON Schema) для извлечения первички через Anthropic Messages.
std::string GeminiPrimaryDocumentAnthropicUserPromptUtf8();

/// Постобработка JSON первички из сырого текста ответа модели (Gemini или Anthropic).
std::string GeminiNormalizePrimaryDocumentJsonFromAssistantText(const std::string& assistant_text_utf8,
                                                                std::string& error_out);

#endif
