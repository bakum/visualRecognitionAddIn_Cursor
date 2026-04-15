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

#include <cstdint>
#include <memory>
#include <string>

#include "Component.h"

/// ПоследняяОшибка (англ. LastErrorCode): 0 успех; 1 пустой PDF/изображение; 4 неверный Base64;
/// 5 неверный тип (blob); 6 неверный тип (строка); 7 пустой ключ API; 8 ошибка Gemini;
/// 9 пустой текст запроса; 10 текст запроса слишком длинный; 99 прочее.
/// Ключ: AIStudioApiKey / КлючAPIAIStudio; модель: GeminiModel / МодельGemini (UTF-8);
/// таймауты: GeminiReceiveTimeoutMs / ТаймаутПолученияGeminiМс, GeminiTotalDeadlineMs / ОбщийДедлайнGeminiМс.
///
/// Локального извлечения текста из PDF нет — только Google Gemini (в т.ч. сканы).
/// Методы РазобратьПервичныйДокументPdf* без «ИИ» в имени — те же вызовы, что и *ИИ (совместимость со старым кодом 1С).
class VisualAddIn : public Component {
public:
    VisualAddIn();

    variant_t ParsePrimaryDocumentPdfAi(variant_t& pdf_blob);
    variant_t ParsePrimaryDocumentPdfAiBase64(variant_t& pdf_base64);

    variant_t ParsePrimaryDocumentImageAi(variant_t& image_blob);
    variant_t ParsePrimaryDocumentImageAiBase64(variant_t& image_base64);

    /// Один аргумент — строка UTF-8 (промпт); ответ модели без принудительного JSON первички.
    variant_t GenerateGeminiText(variant_t& prompt_utf8);

    /// Без параметров: JSON с полями defaultModelId и models[{id,name,notes}] для выбора МодельGemini.
    variant_t GetSupportedGeminiModels();
    /// Без параметров: устанавливает быстрый профиль таймаутов Gemini (получение/общий дедлайн).
    variant_t UseFastGeminiTimeoutsProfile();

protected:
    std::string extensionName() override;

private:
    std::shared_ptr<variant_t> last_error_code_storage_;
    std::shared_ptr<variant_t> last_error_text_storage_;
    std::shared_ptr<variant_t> ai_studio_api_key_storage_;
    std::shared_ptr<variant_t> gemini_model_storage_;
    std::shared_ptr<variant_t> gemini_receive_timeout_ms_storage_;
    std::shared_ptr<variant_t> gemini_total_deadline_ms_storage_;
    std::shared_ptr<variant_t> last_prompt_tokens_storage_;
    std::shared_ptr<variant_t> last_output_tokens_storage_;
    std::shared_ptr<variant_t> last_total_tokens_storage_;
    std::shared_ptr<variant_t> last_usage_json_storage_;

    variant_t ParsePrimaryDocumentGeminiFromBytes(const std::vector<char>& bytes, bool inline_as_pdf);
    void ResetUsageStats();
    void SetUsageStats(int64_t prompt_tokens, int64_t output_tokens, int64_t total_tokens, bool has_usage);
};

#endif // VISUAL_ADDIN_H
