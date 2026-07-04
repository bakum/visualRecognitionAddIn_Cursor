#include "ApiResponseParse.h"

#include <algorithm>
#include <cctype>
#include <limits>
#include <optional>
#include <string>

#include <boost/json.hpp>

std::optional<std::string> ExtractGeminiApiErrorMessage(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& o = v.as_object();
        const auto it = o.find("error");
        if (it == o.end() || !it->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& err = it->value().as_object();
        const auto im = err.find("message");
        if (im == err.end() || !im->value().is_string()) {
            return std::string("Gemini API error");
        }
        return std::string(im->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<std::string> ExtractPromptBlockMessage(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto pc = root.find("promptFeedback");
        if (pc == root.end() || !pc->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& pf = pc->value().as_object();
        const auto pb = pf.find("blockReason");
        if (pb == pf.end()) {
            return std::nullopt;
        }
        if (pb->value().is_string()) {
            return std::string(pb->value().as_string());
        }
        return std::string("blocked");
    } catch (...) {
        return std::nullopt;
    }
}

bool IsTransientGeminiOverload(unsigned long http_status, const std::optional<std::string>& api_err) {
    if (http_status == 429UL || http_status == 503UL) {
        return true;
    }
    if (!api_err.has_value()) {
        return false;
    }
    std::string s = *api_err;
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    return s.find("high demand") != std::string::npos || s.find("try again later") != std::string::npos
        || s.find("resource exhausted") != std::string::npos || s.find("temporar") != std::string::npos;
}

std::optional<std::string> ExtractCandidateText(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto it = root.find("candidates");
        if (it == root.end() || !it->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& cand = it->value().as_array();
        if (cand.empty()) {
            return std::nullopt;
        }
        const boost::json::value& c0 = cand[0];
        if (!c0.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& cobj = c0.as_object();
        const auto cit = cobj.find("content");
        if (cit == cobj.end() || !cit->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& content = cit->value().as_object();
        const auto pit = content.find("parts");
        if (pit == content.end() || !pit->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& parts = pit->value().as_array();
        if (parts.empty() || !parts[0].is_object()) {
            return std::nullopt;
        }
        const boost::json::object& p0 = parts[0].as_object();
        const auto tit = p0.find("text");
        if (tit == p0.end() || !tit->value().is_string()) {
            return std::nullopt;
        }
        return std::string(tit->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<std::string> ExtractCandidateTextAllTextParts(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto it = root.find("candidates");
        if (it == root.end() || !it->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& cand = it->value().as_array();
        if (cand.empty()) {
            return std::nullopt;
        }
        const boost::json::value& c0 = cand[0];
        if (!c0.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& cobj = c0.as_object();
        const auto cit = cobj.find("content");
        if (cit == cobj.end() || !cit->value().is_object()) {
            return std::nullopt;
        }
        const boost::json::object& content = cit->value().as_object();
        const auto pit = content.find("parts");
        if (pit == content.end() || !pit->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& parts = pit->value().as_array();
        std::string merged;
        for (const boost::json::value& pv : parts) {
            if (!pv.is_object()) {
                continue;
            }
            const boost::json::object& po = pv.as_object();
            const auto tit = po.find("text");
            if (tit == po.end() || !tit->value().is_string()) {
                continue;
            }
            merged.append(std::string(tit->value().as_string()));
        }
        if (merged.empty()) {
            return std::nullopt;
        }
        return merged;
    } catch (...) {
        return std::nullopt;
    }
}

void ExtractUsageStats(const std::string& json, GeminiUsageStats* usage_out) {
    if (!usage_out) {
        return;
    }
    *usage_out = GeminiUsageStats{};
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return;
        }
        const boost::json::object& root = v.as_object();
        const auto uit = root.find("usageMetadata");
        if (uit == root.end() || !uit->value().is_object()) {
            return;
        }
        const boost::json::object& usage = uit->value().as_object();
        const auto read_i64 = [&usage](const char* key) -> std::optional<int64_t> {
            const auto it = usage.find(key);
            if (it == usage.end()) {
                return std::nullopt;
            }
            if (it->value().is_int64()) {
                return it->value().as_int64();
            }
            if (it->value().is_uint64()) {
                const uint64_t val = it->value().as_uint64();
                if (val <= static_cast<uint64_t>(std::numeric_limits<int64_t>::max())) {
                    return static_cast<int64_t>(val);
                }
            }
            return std::nullopt;
        };

        const std::optional<int64_t> prompt = read_i64("promptTokenCount");
        const std::optional<int64_t> output = read_i64("candidatesTokenCount");
        const std::optional<int64_t> total = read_i64("totalTokenCount");
        if (prompt.has_value() || output.has_value() || total.has_value()) {
            usage_out->prompt_tokens = prompt.value_or(0);
            usage_out->output_tokens = output.value_or(0);
            usage_out->total_tokens = total.value_or(0);
            usage_out->has_usage = true;
        }
    } catch (...) {
    }
}

std::optional<std::string> ExtractAnthropicApiErrorMessage(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto type_it = root.find("type");
        if (type_it != root.end() && type_it->value().is_string()
            && std::string(type_it->value().as_string()) == "error") {
            const auto err_it = root.find("error");
            if (err_it != root.end() && err_it->value().is_object()) {
                const boost::json::object& err = err_it->value().as_object();
                const auto msg_it = err.find("message");
                if (msg_it != err.end() && msg_it->value().is_string()) {
                    return std::string(msg_it->value().as_string());
                }
            }
        }
        const auto msg_top = root.find("message");
        if (msg_top != root.end() && msg_top->value().is_string()) {
            return std::string(msg_top->value().as_string());
        }
        return std::string("Anthropic API error");
    } catch (...) {
        return std::nullopt;
    }
}

bool IsTransientAnthropicOverload(unsigned long http_status, const std::optional<std::string>& api_err) {
    if (http_status == 429UL || http_status == 503UL || http_status == 529UL) {
        return true;
    }
    if (!api_err.has_value()) {
        return false;
    }
    std::string s = *api_err;
    std::transform(s.begin(), s.end(), s.begin(),
                   [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    return s.find("overload") != std::string::npos || s.find("rate limit") != std::string::npos
        || s.find("try again") != std::string::npos || s.find("temporar") != std::string::npos;
}

void ExtractAnthropicUsage(const std::string& json, GeminiUsageStats* usage_out) {
    if (!usage_out) {
        return;
    }
    *usage_out = GeminiUsageStats{};
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return;
        }
        const boost::json::object& root = v.as_object();
        const auto uit = root.find("usage");
        if (uit == root.end() || !uit->value().is_object()) {
            return;
        }
        const boost::json::object& usage = uit->value().as_object();
        const auto read_i64 = [&usage](const char* key) -> std::optional<int64_t> {
            const auto it = usage.find(key);
            if (it == usage.end()) {
                return std::nullopt;
            }
            if (it->value().is_int64()) {
                return it->value().as_int64();
            }
            if (it->value().is_uint64()) {
                const uint64_t val = it->value().as_uint64();
                if (val <= static_cast<uint64_t>(std::numeric_limits<int64_t>::max())) {
                    return static_cast<int64_t>(val);
                }
            }
            return std::nullopt;
        };
        const std::optional<int64_t> input_tok = read_i64("input_tokens");
        const std::optional<int64_t> output_tok = read_i64("output_tokens");
        if (input_tok.has_value() || output_tok.has_value()) {
            usage_out->prompt_tokens = input_tok.value_or(0);
            usage_out->output_tokens = output_tok.value_or(0);
            usage_out->total_tokens = usage_out->prompt_tokens + usage_out->output_tokens;
            usage_out->has_usage = true;
        }
    } catch (...) {
    }
}

std::optional<std::string> ExtractAssistantText(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& root = v.as_object();
        const auto cit = root.find("content");
        if (cit == root.end() || !cit->value().is_array()) {
            return std::nullopt;
        }
        const boost::json::array& blocks = cit->value().as_array();
        std::string merged;
        for (const boost::json::value& b : blocks) {
            if (!b.is_object()) {
                continue;
            }
            const boost::json::object& bo = b.as_object();
            const auto tit = bo.find("type");
            if (tit == bo.end() || !tit->value().is_string()) {
                continue;
            }
            if (std::string(tit->value().as_string()) != "text") {
                continue;
            }
            const auto tx = bo.find("text");
            if (tx == bo.end() || !tx->value().is_string()) {
                continue;
            }
            merged.append(std::string(tx->value().as_string()));
        }
        if (merged.empty()) {
            return std::nullopt;
        }
        return merged;
    } catch (...) {
        return std::nullopt;
    }
}

std::optional<std::string> ExtractAnthropicStopReason(const std::string& json) {
    try {
        const boost::json::value v = boost::json::parse(json);
        if (!v.is_object()) {
            return std::nullopt;
        }
        const boost::json::object& o = v.as_object();
        const auto it = o.find("stop_reason");
        if (it == o.end() || !it->value().is_string()) {
            return std::nullopt;
        }
        return std::string(it->value().as_string());
    } catch (...) {
        return std::nullopt;
    }
}
