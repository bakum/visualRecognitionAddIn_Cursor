#include "ModelResolve.h"

#include <cctype>
#include <string>

#include "AnthropicMessagesClient.h"
#include "GeminiPdfClient.h"
#include "StringUtils.h"

namespace {

std::string BuildModelAliasKey(std::string_view value) {
    std::string key;
    key.reserve(value.size());
    for (unsigned char uc : value) {
        if (std::isalnum(uc) == 0) {
            continue;
        }
        key.push_back(static_cast<char>(std::tolower(uc)));
    }
    return key;
}

} // namespace

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

std::string ResolveGeminiModelAlias(std::string_view model_id_utf8) {
    std::string model(model_id_utf8);
    TrimAsciiInPlace(model);
    if (model.empty()) {
        return std::string(kGeminiDefaultModelId);
    }

    const std::string key = BuildModelAliasKey(model);
    if (key == "gemini3pro" || key == "gemini31pro") {
        return "gemini-3.1-pro-preview";
    }
    if (key == "gemini31flashlitepreview") {
        return "gemini-3.1-flash-lite";
    }
    if (key == "gemini3flashlite" || key == "gemini31flashlite") {
        return "gemini-3.1-flash-lite";
    }
    return model;
}

bool IsValidAnthropicModelId(std::string_view s) {
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

std::string ResolveAnthropicModelAlias(std::string_view model_id_utf8) {
    std::string model(model_id_utf8);
    TrimAsciiInPlace(model);
    if (model.empty()) {
        return std::string(kAnthropicDefaultModelId);
    }
    const std::string key = BuildModelAliasKey(model);
    if (key == "claudesonnet45" || key == "sonnet45" || key == "claude45sonnet") {
        return "claude-sonnet-4-5-20250929";
    }
    if (key == "claudehaiku45" || key == "haiku45") {
        return "claude-haiku-4-5-20251001";
    }
    if (key == "claude35sonnet" || key == "sonnet35") {
        return "claude-3-5-sonnet-20241022";
    }
    if (key == "claude35haiku" || key == "haiku35") {
        return "claude-3-5-haiku-20241022";
    }
    if (key == "claude3opus" || key == "opus3") {
        return "claude-3-opus-20240229";
    }
    return model;
}
