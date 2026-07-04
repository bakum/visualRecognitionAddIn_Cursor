#pragma once

#include <boost/json.hpp>
#include <string>

inline boost::json::object ParseJsonObject(const std::string& json_utf8) {
    const boost::json::value v = boost::json::parse(json_utf8);
    return v.as_object();
}

inline std::string JsonStringField(const boost::json::object& o, const char* key) {
    const auto it = o.find(key);
    if (it == o.end() || !it->value().is_string()) {
        return {};
    }
    return std::string(it->value().as_string());
}

inline std::string MinimalPrimaryDocSkeleton(std::string document_title,
                                             std::string document_type,
                                             std::string document_date = "",
                                             std::string contract_details = "") {
    boost::json::object root;
    boost::json::object counterparty;
    counterparty["supplier"] = "";
    counterparty["buyer"] = "";
    counterparty["supplierInn"] = "";
    counterparty["buyerInn"] = "";
    counterparty["supplierKpp"] = "";
    counterparty["buyerKpp"] = "";
    counterparty["supplierOkpo"] = "";
    counterparty["buyerOkpo"] = "";
    root["counterparty"] = counterparty;

    boost::json::object contract;
    contract["numberAndDetails"] = contract_details;
    contract["date"] = "";
    root["contract"] = contract;

    root["invoiceNumber"] = "";
    root["documentDate"] = document_date;
    root["documentTitle"] = document_title;
    root["documentType"] = document_type;
    root["bankDetails"] = boost::json::object();
    root["lineItems"] = boost::json::array();
    return boost::json::serialize(root);
}
