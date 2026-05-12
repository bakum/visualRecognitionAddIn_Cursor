#ifndef ANTHROPIC_MESSAGES_CLIENT_H
#define ANTHROPIC_MESSAGES_CLIENT_H

#include <string>
#include <string_view>

#include "GeminiPdfClient.h"

/// Идентификатор модели по умолчанию (если из 1С передана пустая строка).
inline constexpr char kAnthropicDefaultModelId[] = "claude-sonnet-4-20250514";

/// Текстовый запрос к Anthropic Messages API; ответ — строка UTF-8 (контент ассистента).
/// Ключ: свойства аддина AnthropicApiKey / КлючAPIAnthropic; модель: AnthropicModel / МодельAnthropic.
/// Таймауты: те же, что для Gemini (GeminiReceiveTimeoutMs, GeminiTotalDeadlineMs).
/// max_tokens: AnthropicMaxOutputTokens / МаксТокеновВыводаAnthropic.
std::string AnthropicGeneratePlainText(const std::string& api_key_utf8,
                                       const std::string& model_id_utf8,
                                       int max_output_tokens,
                                       const std::string& user_text_utf8,
                                       std::string& error_out,
                                       GeminiUsageStats* usage_out = nullptr,
                                       const GeminiHttpTimeouts* timeouts = nullptr,
                                       std::string* raw_response_out = nullptr);

/// Каталог рекомендуемых моделей (UTF-8 JSON): defaultModelId, models[{id,name,notes}]. Ключ API не нужен.
std::string AnthropicSupportedModelsCatalogJson();

/// PDF или растровое изображение → JSON первички (та же постобработка, что у Gemini).
std::string AnthropicExtractPrimaryDocumentJson(const std::string& api_key_utf8,
                                                 const std::string& model_id_utf8,
                                                 std::string_view mime_type,
                                                 const std::vector<char>& bytes,
                                                 int max_output_tokens,
                                                 std::string& error_out,
                                                 GeminiUsageStats* usage_out = nullptr,
                                                 const GeminiHttpTimeouts* timeouts = nullptr,
                                                 std::string* raw_response_out = nullptr);

#endif
