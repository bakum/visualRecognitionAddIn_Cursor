#include <catch2/catch_test_macros.hpp>

#include "GeminiPdfClient.h"

#include "test_helpers.h"

TEST_CASE("Infer documentType from documentTitle", "[normalize]") {
    const std::string input = MinimalPrimaryDocSkeleton("Invoice No 123", u8"Неопределено");
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText(input, err);
    REQUIRE(err.empty());
    const boost::json::object root = ParseJsonObject(out);
    REQUIRE(JsonStringField(root, "documentType") == u8"Счет");
}

TEST_CASE("Infer incoming waybill documentType", "[normalize]") {
    const std::string title = u8"приходная накладная 45";
    const std::string input = MinimalPrimaryDocSkeleton(title, u8"Неопределено");
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText(input, err);
    REQUIRE(err.empty());
    const boost::json::object root = ParseJsonObject(out);
    REQUIRE(JsonStringField(root, "documentType") == u8"ПриходнаяНакладная");
}

TEST_CASE("Normalize documentDate YYYYMMDD", "[normalize]") {
    const std::string input = MinimalPrimaryDocSkeleton(u8"Счет", u8"Счет", "20260315");
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText(input, err);
    REQUIRE(err.empty());
    const boost::json::object root = ParseJsonObject(out);
    REQUIRE(JsonStringField(root, "documentDate") == "2026/03/15");
}

TEST_CASE("Strip markdown JSON fence from assistant text", "[normalize]") {
    const std::string inner = MinimalPrimaryDocSkeleton(u8"Invoice", u8"Счет");
    const std::string fenced = std::string("```json\n") + inner + "\n```";
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText(fenced, err);
    REQUIRE(err.empty());
    REQUIRE(ParseJsonObject(out).size() > 0U);
}

TEST_CASE("Normalize contract numberAndDetails", "[normalize]") {
    const std::string input = MinimalPrimaryDocSkeleton(u8"Счет", u8"Счет", "", "N 123/45");
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText(input, err);
    REQUIRE(err.empty());
    const boost::json::object root = ParseJsonObject(out);
    const auto contract_it = root.find("contract");
    REQUIRE(contract_it != root.end());
    REQUIRE(contract_it->value().is_object());
    const boost::json::object& contract = contract_it->value().as_object();
    const std::string details = JsonStringField(contract, "numberAndDetails");
    REQUIRE(details == "123/45");
}

TEST_CASE("Sanitize line item name non-printable characters", "[normalize]") {
    boost::json::object root = ParseJsonObject(MinimalPrimaryDocSkeleton(u8"Счет", u8"Счет"));
    boost::json::object line;
    line["name"] = std::string("Item\u0001\u0002  name");
    line["sku"] = "";
    line["barcode"] = "";
    line["quantity"] = "1";
    line["unit"] = "шт";
    line["price"] = "10";
    line["priceVatType"] = "";
    line["vatRate"] = "";
    line["amount"] = "10";
    boost::json::array items;
    items.push_back(line);
    root["lineItems"] = items;

    std::string err;
    const std::string out =
        GeminiNormalizePrimaryDocumentJsonFromAssistantText(boost::json::serialize(root), err);
    REQUIRE(err.empty());
    const boost::json::object parsed = ParseJsonObject(out);
    const auto li_it = parsed.find("lineItems");
    REQUIRE(li_it != parsed.end());
    const boost::json::array& arr = li_it->value().as_array();
    REQUIRE(arr.size() == 1U);
    const boost::json::object& item = arr[0].as_object();
    REQUIRE(JsonStringField(item, "name") == "Item name");
}

TEST_CASE("Empty assistant text returns error", "[normalize]") {
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText("   ", err);
    REQUIRE(out.empty());
    REQUIRE(err == "Empty model JSON body");
}

TEST_CASE("Invalid JSON returns parse error", "[normalize]") {
    std::string err;
    const std::string out = GeminiNormalizePrimaryDocumentJsonFromAssistantText("{not json", err);
    REQUIRE(out.empty());
    REQUIRE(err.find("Model JSON parse error") != std::string::npos);
}
