#include <catch2/catch_test_macros.hpp>

#include <stdexcept>
#include <string>
#include <vector>

#include <boost/json.hpp>

#include "VisualAddIn.h"

TEST_CASE("VisualAddIn EncodeToBase64 / DecodeFromBase64 round-trip", "[visual_addin]") {
    VisualAddIn addin;
    variant_t blob = std::vector<char>{'A', 'B', 'C'};
    variant_t encoded = addin.EncodeToBase64(blob);
    REQUIRE(std::holds_alternative<std::string>(encoded));

    variant_t decoded = addin.DecodeFromBase64(encoded);
    REQUIRE(std::holds_alternative<std::vector<char>>(decoded));
    REQUIRE(std::get<std::vector<char>>(decoded) == std::get<std::vector<char>>(blob));
}

TEST_CASE("VisualAddIn DecodeFromBase64 invalid input", "[visual_addin]") {
    VisualAddIn addin;
    variant_t bad = std::string("***");
    variant_t result = addin.DecodeFromBase64(bad);
    REQUIRE(std::holds_alternative<std::string>(result));
    REQUIRE(std::get<std::string>(result).find("Invalid Base64 string") != std::string::npos);
}

TEST_CASE("VisualAddIn EncryptTextSimple / DecryptTextSimple round-trip", "[visual_addin]") {
    VisualAddIn addin;
    variant_t plain = std::string("secret-value-123");
    variant_t encrypted = addin.EncryptTextSimple(plain);
    REQUIRE(std::holds_alternative<std::string>(encrypted));

    variant_t decrypted = addin.DecryptTextSimple(encrypted);
    REQUIRE(std::holds_alternative<std::string>(decrypted));
    REQUIRE(std::get<std::string>(decrypted) == std::get<std::string>(plain));
}

TEST_CASE("VisualAddIn EncodeKey / DecodeKey round-trip", "[visual_addin]") {
    VisualAddIn addin;
    variant_t key = std::string("my-api-key-2026");
    variant_t packed = addin.EncodeKey(key);
    REQUIRE(std::holds_alternative<std::string>(packed));

    variant_t restored = addin.DecodeKey(packed);
    REQUIRE(std::holds_alternative<std::string>(restored));
    REQUIRE(std::get<std::string>(restored) == std::get<std::string>(key));
}

TEST_CASE("VisualAddIn wrong argument types throw", "[visual_addin]") {
    VisualAddIn addin;
    variant_t not_blob = std::string("text");
    REQUIRE_THROWS_AS(addin.EncodeToBase64(not_blob), std::runtime_error);
    REQUIRE_THROWS_AS(addin.ParsePrimaryDocumentPdfAi(not_blob), std::runtime_error);
}

TEST_CASE("VisualAddIn ParsePrimaryDocumentPdf empty blob returns error JSON", "[visual_addin]") {
    VisualAddIn addin;
    variant_t empty_pdf = std::vector<char>{};
    variant_t result = addin.ParsePrimaryDocumentPdfAi(empty_pdf);
    REQUIRE(std::holds_alternative<std::string>(result));
    REQUIRE(std::get<std::string>(result).find("Empty PDF data") != std::string::npos);
}

TEST_CASE("VisualAddIn GetSupportedGeminiModels returns catalog JSON", "[visual_addin]") {
    VisualAddIn addin;
    variant_t catalog = addin.GetSupportedGeminiModels();
    REQUIRE(std::holds_alternative<std::string>(catalog));
    const boost::json::value v = boost::json::parse(std::get<std::string>(catalog));
    REQUIRE(v.is_object());
    REQUIRE(v.as_object().if_contains("defaultModelId") != nullptr);
}

TEST_CASE("VisualAddIn UseFastGeminiTimeoutsProfile", "[visual_addin]") {
    VisualAddIn addin;
    variant_t profile = addin.UseFastGeminiTimeoutsProfile();
    REQUIRE(std::holds_alternative<std::string>(profile));
    const boost::json::object o = boost::json::parse(std::get<std::string>(profile)).as_object();
    REQUIRE(std::string(o.at("profile").as_string()) == "fast");
    REQUIRE(o.at("receiveTimeoutMs").as_int64() == 20000);
    REQUIRE(o.at("totalDeadlineMs").as_int64() == 30000);
}

TEST_CASE("VisualAddIn GenerateGeminiText without API key", "[visual_addin]") {
    VisualAddIn addin;
    variant_t prompt = std::string("hello");
    variant_t result = addin.GenerateGeminiText(prompt);
    REQUIRE(std::holds_alternative<std::string>(result));
    REQUIRE(std::get<std::string>(result).find("Gemini API key is empty") != std::string::npos);
}
