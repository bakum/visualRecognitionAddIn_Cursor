#ifndef ADDIN_UTILS_H
#define ADDIN_UTILS_H

#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "Component.h"
#include "GeminiPdfClient.h"

namespace addin {

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

std::optional<std::vector<char>> DecodeBase64(std::string_view in);
std::string EncodeBase64(const std::vector<char>& in);

std::string EncryptUtf8ToHex(std::string_view plain_utf8);
std::optional<std::string> DecryptHexToUtf8(std::string_view encrypted_hex);

std::optional<std::string> ExtractErrorFromResultJson(const std::string& json_utf8);

void SetLastErrorPair(const std::shared_ptr<variant_t>& code_storage,
                      const std::shared_ptr<variant_t>& text_storage,
                      int32_t code,
                      std::string text_utf8);
void ClearLastErrorPair(const std::shared_ptr<variant_t>& code_storage,
                        const std::shared_ptr<variant_t>& text_storage);

std::pair<int32_t, std::string> MapPipelineErrorToCodeAndText(const std::string& error_field_utf8);

void SyncLastErrorFromParseResult(const std::shared_ptr<variant_t>& code_storage,
                                  const std::shared_ptr<variant_t>& text_storage,
                                  const variant_t& r);

std::string GetTrimmedUtf8Property(const std::shared_ptr<variant_t>& storage);

std::string AlphanumericAsciiLowerKey(std::string_view s);
bool LooksLikeAnthropicModelId(const std::string& model_id);
bool LooksLikeGeminiModelId(const std::string& model_id);

void StripUtf8BomAndNbspEdgesInPlace(std::string& v);
bool AiBackendIsAnthropic(const std::shared_ptr<variant_t>& ai_backend_storage);

int GetIntPropertyOrDefault(const std::shared_ptr<variant_t>& storage,
                            int default_value,
                            int min_value,
                            int max_value);

variant_t JsonErrorObjectUtf8(const std::string& error_utf8);
std::string UsageJsonFromStats(const GeminiUsageStats& usage);
std::string MakeUtf8Preview(const std::string& s, size_t max_bytes);

std::optional<std::string> DetectRasterImageMimeForAi(const std::vector<char>& b);

} // namespace addin

#endif
