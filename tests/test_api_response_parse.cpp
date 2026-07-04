#include <catch2/catch_test_macros.hpp>

#include "ApiResponseParse.h"

TEST_CASE("ExtractGeminiApiErrorMessage", "[api_response]") {
    const std::string json = R"({"error":{"message":"API key not valid","status":"INVALID_ARGUMENT"}})";
    REQUIRE(ExtractGeminiApiErrorMessage(json) == "API key not valid");
    REQUIRE_FALSE(ExtractGeminiApiErrorMessage(R"({"candidates":[]})").has_value());
}

TEST_CASE("ExtractPromptBlockMessage", "[api_response]") {
    const std::string json = R"({"promptFeedback":{"blockReason":"SAFETY"}})";
    REQUIRE(ExtractPromptBlockMessage(json) == "SAFETY");
}

TEST_CASE("ExtractCandidateText and all parts", "[api_response]") {
    const std::string single = R"({
        "candidates":[{"content":{"parts":[{"text":"{\"invoiceNumber\":\"1\"}"}]}}]
    })";
    REQUIRE(ExtractCandidateText(single) == R"({"invoiceNumber":"1"})");

    const std::string multi = R"({
        "candidates":[{"content":{"parts":[
            {"text":"line1\n"},
            {"text":"line2"}
        ]}}]
    })";
    REQUIRE(ExtractCandidateTextAllTextParts(multi) == "line1\nline2");
}

TEST_CASE("ExtractUsageStats from Gemini response", "[api_response]") {
    const std::string json = R"({"usageMetadata":{"promptTokenCount":100,"candidatesTokenCount":50,"totalTokenCount":150}})";
    GeminiUsageStats usage;
    ExtractUsageStats(json, &usage);
    REQUIRE(usage.has_usage);
    REQUIRE(usage.prompt_tokens == 100);
    REQUIRE(usage.output_tokens == 50);
    REQUIRE(usage.total_tokens == 150);
}

TEST_CASE("IsTransientGeminiOverload", "[api_response]") {
    REQUIRE(IsTransientGeminiOverload(503UL, std::nullopt));
    REQUIRE(IsTransientGeminiOverload(400UL, std::string("Resource exhausted, try again later")));
    REQUIRE_FALSE(IsTransientGeminiOverload(400UL, std::string("invalid api key")));
}

TEST_CASE("ExtractAnthropicApiErrorMessage", "[api_response]") {
    const std::string json =
        R"({"type":"error","error":{"type":"authentication_error","message":"invalid x-api-key"}})";
    REQUIRE(ExtractAnthropicApiErrorMessage(json) == "invalid x-api-key");
}

TEST_CASE("ExtractAssistantText from Anthropic response", "[api_response]") {
    const std::string json = R"({
        "content":[
            {"type":"text","text":"Hello"},
            {"type":"tool_use","id":"x","name":"y","input":{}}
        ]
    })";
    REQUIRE(ExtractAssistantText(json) == "Hello");
}

TEST_CASE("ExtractAnthropicUsage", "[api_response]") {
    const std::string json = R"({"usage":{"input_tokens":11,"output_tokens":22}})";
    GeminiUsageStats usage;
    ExtractAnthropicUsage(json, &usage);
    REQUIRE(usage.has_usage);
    REQUIRE(usage.prompt_tokens == 11);
    REQUIRE(usage.output_tokens == 22);
    REQUIRE(usage.total_tokens == 33);
}

TEST_CASE("ExtractAnthropicStopReason", "[api_response]") {
    const std::string json = R"({"stop_reason":"max_tokens","content":[]})";
    REQUIRE(ExtractAnthropicStopReason(json) == "max_tokens");
}

TEST_CASE("IsTransientAnthropicOverload", "[api_response]") {
    REQUIRE(IsTransientAnthropicOverload(529UL, std::nullopt));
    REQUIRE(IsTransientAnthropicOverload(400UL, std::string("Overloaded")));
    REQUIRE_FALSE(IsTransientAnthropicOverload(400UL, std::string("invalid request")));
}
