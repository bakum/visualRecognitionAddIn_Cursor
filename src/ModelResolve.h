#ifndef MODEL_RESOLVE_H
#define MODEL_RESOLVE_H

#include <string>
#include <string_view>

bool IsValidGeminiModelId(std::string_view s);
std::string ResolveGeminiModelAlias(std::string_view model_id_utf8);

bool IsValidAnthropicModelId(std::string_view s);
std::string ResolveAnthropicModelAlias(std::string_view model_id_utf8);

#endif
