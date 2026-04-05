#ifndef GEMINI_PDF_CLIENT_H
#define GEMINI_PDF_CLIENT_H

#include <string>
#include <vector>

/// Идентификатор модели по умолчанию (если из 1С передана пустая строка).
inline constexpr char kGeminiDefaultModelId[] = "gemini-2.5-flash";

/// PDF (application/pdf) → Gemini; структура как у локального разбора первички.
/// model_id_utf8 — имя модели API; пусто после trim — kGeminiDefaultModelId.
std::string GeminiExtractPrimaryDocumentJson(const std::string& api_key_utf8,
                                             const std::string& model_id_utf8,
                                             const std::vector<char>& pdf_bytes,
                                             std::string& error_out);

/// Растровое изображение (JPEG/PNG/GIF/WebP/BMP/TIFF по сигнатуре). PDF не принимается — для PDF используйте PdfИИ.
std::string GeminiExtractPrimaryDocumentJsonFromImageBytes(const std::string& api_key_utf8,
                                                           const std::string& model_id_utf8,
                                                           const std::vector<char>& image_bytes,
                                                           std::string& error_out);

#endif
