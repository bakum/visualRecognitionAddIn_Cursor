#include <catch2/catch_test_macros.hpp>

#include "AnthropicMessagesClient.h"
#include "GeminiPdfClient.h"
#include "ModelResolve.h"

TEST_CASE("ResolveGeminiModelAlias", "[model_resolve]") {
    REQUIRE(ResolveGeminiModelAlias("") == kGeminiDefaultModelId);
    REQUIRE(ResolveGeminiModelAlias("  ") == kGeminiDefaultModelId);
    REQUIRE(ResolveGeminiModelAlias("gemini31flashlite") == "gemini-3.1-flash-lite");
    REQUIRE(ResolveGeminiModelAlias("gemini-3.1-pro-preview") == "gemini-3.1-pro-preview");
}

TEST_CASE("IsValidGeminiModelId", "[model_resolve]") {
    REQUIRE(IsValidGeminiModelId("gemini-3-flash-preview"));
    REQUIRE(IsValidGeminiModelId("model.name_v2"));
    REQUIRE_FALSE(IsValidGeminiModelId(""));
    REQUIRE_FALSE(IsValidGeminiModelId("model/with/slash"));
    REQUIRE_FALSE(IsValidGeminiModelId(std::string(161, 'a')));
}

TEST_CASE("ResolveAnthropicModelAlias", "[model_resolve]") {
    REQUIRE(ResolveAnthropicModelAlias("") == kAnthropicDefaultModelId);
    REQUIRE(ResolveAnthropicModelAlias("haiku45") == "claude-haiku-4-5-20251001");
    REQUIRE(ResolveAnthropicModelAlias("sonnet45") == "claude-sonnet-4-5-20250929");
    REQUIRE(ResolveAnthropicModelAlias("claude-sonnet-4-5-20250929") == "claude-sonnet-4-5-20250929");
}

TEST_CASE("IsValidAnthropicModelId", "[model_resolve]") {
    REQUIRE(IsValidAnthropicModelId("claude-haiku-4-5-20251001"));
    REQUIRE_FALSE(IsValidAnthropicModelId("bad model"));
    REQUIRE_FALSE(IsValidAnthropicModelId("claude?test"));
}

TEST_CASE("GeminiSupportedModelsCatalogJson structure", "[model_resolve]") {
    const std::string json = GeminiSupportedModelsCatalogJson();
    REQUIRE(json.find("defaultModelId") != std::string::npos);
    REQUIRE(json.find("models") != std::string::npos);
    REQUIRE(json.find(kGeminiDefaultModelId) != std::string::npos);
}

TEST_CASE("AnthropicSupportedModelsCatalogJson structure", "[model_resolve]") {
    const std::string json = AnthropicSupportedModelsCatalogJson();
    REQUIRE(json.find("defaultModelId") != std::string::npos);
    REQUIRE(json.find("models") != std::string::npos);
    REQUIRE(json.find(kAnthropicDefaultModelId) != std::string::npos);
}
