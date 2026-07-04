#ifndef API_RESPONSE_PARSE_H
#define API_RESPONSE_PARSE_H

#include <cstdint>
#include <optional>
#include <string>

#include "GeminiPdfClient.h"

std::optional<std::string> ExtractGeminiApiErrorMessage(const std::string& json);
std::optional<std::string> ExtractPromptBlockMessage(const std::string& json);
bool IsTransientGeminiOverload(unsigned long http_status, const std::optional<std::string>& api_err);
std::optional<std::string> ExtractCandidateText(const std::string& json);
std::optional<std::string> ExtractCandidateTextAllTextParts(const std::string& json);
void ExtractUsageStats(const std::string& json, GeminiUsageStats* usage_out);

std::optional<std::string> ExtractAnthropicApiErrorMessage(const std::string& json);
bool IsTransientAnthropicOverload(unsigned long http_status, const std::optional<std::string>& api_err);
void ExtractAnthropicUsage(const std::string& json, GeminiUsageStats* usage_out);
std::optional<std::string> ExtractAssistantText(const std::string& json);
std::optional<std::string> ExtractAnthropicStopReason(const std::string& json);

#endif
