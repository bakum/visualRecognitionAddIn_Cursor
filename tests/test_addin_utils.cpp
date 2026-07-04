#include <catch2/catch_test_macros.hpp>

#include <string>
#include <vector>

#include "AddInUtils.h"

using namespace addin;

TEST_CASE("Base64 round-trip and edge cases", "[addin_utils]") {
    const std::vector<char> bytes = {'h', 'e', 'l', 'l', 'o', '\0', char(0xFF)};
    const std::string encoded = EncodeBase64(bytes);
    REQUIRE(encoded == "aGVsbG8A/w==");

    const auto decoded = DecodeBase64(encoded);
    REQUIRE(decoded.has_value());
    REQUIRE(*decoded == bytes);

    REQUIRE(DecodeBase64("aGVsbG8").has_value());
    REQUIRE(DecodeBase64("aGVsbG 8=").has_value());
    REQUIRE_FALSE(DecodeBase64("***").has_value());
    REQUIRE_FALSE(DecodeBase64("ab!").has_value());
    REQUIRE_FALSE(DecodeBase64("abcde").has_value());
}

TEST_CASE("Simple cipher round-trip", "[addin_utils]") {
    const std::string plain = u8"Ключ API 2026! test";
    const std::string encrypted = EncryptUtf8ToHex(plain);
    REQUIRE(encrypted.size() == plain.size() * 2U);

    const auto decrypted = DecryptHexToUtf8(encrypted);
    REQUIRE(decrypted.has_value());
    REQUIRE(*decrypted == plain);

    REQUIRE_FALSE(DecryptHexToUtf8("ABC").has_value());
    REQUIRE_FALSE(DecryptHexToUtf8("GG").has_value());
}

TEST_CASE("MapPipelineErrorToCodeAndText known errors", "[addin_utils]") {
    const auto empty_pdf = MapPipelineErrorToCodeAndText("Empty PDF data");
    REQUIRE(empty_pdf.first == kErrEmptyPdf);
    REQUIRE(empty_pdf.second == u8"Пустые данные PDF.");

    const auto gemini_key = MapPipelineErrorToCodeAndText("Gemini API key is empty");
    REQUIRE(gemini_key.first == kErrMissingApiKey);

    const auto anthropic_key = MapPipelineErrorToCodeAndText("Anthropic API key is empty");
    REQUIRE(anthropic_key.first == kErrMissingApiKey);

    const auto mixed = MapPipelineErrorToCodeAndText("Anthropic model id used with Gemini provider");
    REQUIRE(mixed.first == kErrAiFailure);

    const auto unknown = MapPipelineErrorToCodeAndText("Some custom failure");
    REQUIRE(unknown.first == kErrUnknown);
    REQUIRE(unknown.second == "Some custom failure");
}

TEST_CASE("ExtractErrorFromResultJson", "[addin_utils]") {
    REQUIRE(ExtractErrorFromResultJson(R"({"error":"Invalid Base64 string"})") == "Invalid Base64 string");
    REQUIRE_FALSE(ExtractErrorFromResultJson(R"({"ok":true})").has_value());
    REQUIRE_FALSE(ExtractErrorFromResultJson("not json").has_value());
}

TEST_CASE("AiBackendIsAnthropic provider detection", "[addin_utils]") {
    const auto make_storage = [](std::string v) {
        return std::make_shared<variant_t>(std::move(v));
    };

    REQUIRE_FALSE(AiBackendIsAnthropic(make_storage("gemini")));
    REQUIRE(AiBackendIsAnthropic(make_storage("anthropic")));
    REQUIRE(AiBackendIsAnthropic(make_storage("claude-3")));
    REQUIRE(AiBackendIsAnthropic(make_storage(u8"антропик")));
    REQUIRE(AiBackendIsAnthropic(make_storage(u8"Клод")));

    std::string with_bom = std::string("\xEF\xBB\xBF", 3) + "anthropic";
    REQUIRE(AiBackendIsAnthropic(make_storage(with_bom)));

    REQUIRE_FALSE(AiBackendIsAnthropic(std::make_shared<variant_t>(int32_t{1})));
}

TEST_CASE("LooksLike model id heuristics", "[addin_utils]") {
    REQUIRE(LooksLikeAnthropicModelId("claude-haiku-4-5"));
    REQUIRE(LooksLikeAnthropicModelId("  haiku45  "));
    REQUIRE(LooksLikeGeminiModelId("gemini-3-flash-preview"));
    REQUIRE(LooksLikeGeminiModelId("Gemini31FlashLite"));
    REQUIRE_FALSE(LooksLikeGeminiModelId("claude-haiku-4-5"));
    REQUIRE_FALSE(LooksLikeAnthropicModelId("gemini-3-flash"));
}

TEST_CASE("DetectRasterImageMimeForAi signatures", "[addin_utils]") {
    const std::vector<char> jpeg = {char(0xFF), char(0xD8), char(0xFF), 0};
    REQUIRE(DetectRasterImageMimeForAi(jpeg) == "image/jpeg");

    const std::vector<char> png = {char(0x89), 'P', 'N', 'G', char(0x0D), char(0x0A), char(0x1A), char(0x0A)};
    REQUIRE(DetectRasterImageMimeForAi(png) == "image/png");

    const std::vector<char> pdf = {'%', 'P', 'D', 'F'};
    REQUIRE_FALSE(DetectRasterImageMimeForAi(pdf).has_value());
}

TEST_CASE("GetIntPropertyOrDefault clamps values", "[addin_utils]") {
    auto storage = std::make_shared<variant_t>(int32_t{999999});
    REQUIRE(GetIntPropertyOrDefault(storage, 100, 1000, 300000) == 300000);

    *storage = int32_t{50};
    REQUIRE(GetIntPropertyOrDefault(storage, 100, 1000, 300000) == 1000);

    *storage = double{12345.7};
    REQUIRE(GetIntPropertyOrDefault(storage, 100, 1000, 300000) == 12345);
}

TEST_CASE("UsageJsonFromStats and MakeUtf8Preview", "[addin_utils]") {
    GeminiUsageStats empty;
    REQUIRE(UsageJsonFromStats(empty) == "{}");

    GeminiUsageStats usage;
    usage.prompt_tokens = 10;
    usage.output_tokens = 20;
    usage.total_tokens = 30;
    usage.has_usage = true;
    const std::string usage_json = UsageJsonFromStats(usage);
    REQUIRE(usage_json.find("10") != std::string::npos);
    REQUIRE(usage_json.find("20") != std::string::npos);

    std::string long_text(5000, 'x');
    const std::string preview = MakeUtf8Preview(long_text, 4096);
    REQUIRE(preview.size() == 4096U + std::string("...(truncated)").size());
}
