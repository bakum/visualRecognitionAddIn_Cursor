#ifndef STRING_UTILS_H
#define STRING_UTILS_H

#include <cctype>
#include <string>

inline void TrimAsciiInPlace(std::string& s) {
    while (!s.empty() && (static_cast<unsigned char>(s.front()) <= ' ')) {
        s.erase(s.begin());
    }
    while (!s.empty() && (static_cast<unsigned char>(s.back()) <= ' ')) {
        s.pop_back();
    }
}

#endif
