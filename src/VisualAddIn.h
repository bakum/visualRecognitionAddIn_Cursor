/*
 *  Modern Native AddIn
 *  Copyright (C) 2018  Infactum
 *
 *  This program is free software: you can redistribute it and/or modify
 *  it under the terms of the GNU Affero General Public License as
 *  published by the Free Software Foundation, either version 3 of the
 *  License, or (at your option) any later version.
 *
 *  This program is distributed in the hope that it will be useful,
 *  but WITHOUT ANY WARRANTY; without even the implied warranty of
 *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *  GNU Affero General Public License for more details.
 *
 *  You should have received a copy of the GNU Affero General Public License
 *  along with this program.  If not, see <https://www.gnu.org/licenses/>.
 *
 */

#ifndef VISUAL_ADDIN_H
#define VISUAL_ADDIN_H

#include <memory>
#include <string>

#include "Component.h"

/// ПоследняяОшибка (англ. LastErrorCode) — целое: 0 успех; 1 пустой PDF; 2 нет текста/скан;
/// 3 PDF отключён при сборке; 4 неверный Base64; 5 неверный тип аргумента (blob); 6 неверный тип (строка);
/// 7 пустой ключ API; 8 ошибка Gemini (сеть/API); 99 прочая ошибка (текст в ТекстПоследнейОшибки / LastErrorText).
/// Ключ API AI Studio: AIStudioApiKey / КлючAPIAIStudio; модель Gemini: GeminiModel / МодельGemini (UTF-8, чтение и запись).
class VisualAddIn : public Component {
public:
    VisualAddIn();

    variant_t ParsePrimaryDocumentPdf(variant_t& pdf_blob);

    /// Те саме, що РазобратьПервичныйДокументPdf, але PDF передається як Base64 (рядок UTF-8).
    variant_t ParsePrimaryDocumentPdfBase64(variant_t& pdf_base64);

    /// Плоский текст PDF (UTF-8), без разбора первички. Коды ошибок — те же, что у РазобратьПервичныйДокументPdf.
    variant_t ExtractPdfPlainText(variant_t& pdf_blob);
    variant_t ExtractPdfPlainTextBase64(variant_t& pdf_base64);

    /// Структурирование через Google Gemini (PDF со сканом и/или текстовым слоем); ключ — КлючAPIAIStudio / AIStudioApiKey. JSON как у РазобратьПервичныйДокументPdf.
    variant_t ParsePrimaryDocumentPdfAi(variant_t& pdf_blob);
    variant_t ParsePrimaryDocumentPdfAiBase64(variant_t& pdf_base64);

    /// То же через Gemini, но вход — одно растровое изображение (JPEG/PNG/GIF/WebP/BMP/TIFF); PDF не допускается.
    variant_t ParsePrimaryDocumentImageAi(variant_t& image_blob);
    variant_t ParsePrimaryDocumentImageAiBase64(variant_t& image_base64);

protected:
    std::string extensionName() override;

private:
    std::shared_ptr<variant_t> last_error_code_storage_;
    /// UTF-8: пусто при коде 0; иначе текст для пользователя.
    std::shared_ptr<variant_t> last_error_text_storage_;
    /// UTF-8: ключ API Google AI Studio (aistudio.google.com); задаётся из 1С перед вызовами, требующими ключа.
    std::shared_ptr<variant_t> ai_studio_api_key_storage_;
    /// UTF-8: id модели Gemini (например gemini-2.0-flash). Пустая строка — внутри подставляется значение по умолчанию.
    std::shared_ptr<variant_t> gemini_model_storage_;

    variant_t ParsePrimaryDocumentGeminiFromBytes(const std::vector<char>& bytes, bool inline_as_pdf);
};

#endif // VISUAL_ADDIN_H
