#ifndef PDF_TEXT_EXTRACTOR_H
#define PDF_TEXT_EXTRACTOR_H

#include <cstdint>
#include <string>
#include <vector>

// Extracts plain text from PDF bytes (text layer). Thread-safe internally.
std::wstring ExtractPdfTextW(const std::vector<char>& pdf_bytes);

#endif
