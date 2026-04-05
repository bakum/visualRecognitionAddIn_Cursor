#ifndef INVOICE_TEXT_PARSER_H
#define INVOICE_TEXT_PARSER_H

#include <string>

// Евристичний розбір текстового шару українських первинок (рахунок, рахунок-фактура, акт тощо).
// Повертає компактний JSON (UTF-8, серіалізація Boost.JSON).
std::string ParseInvoiceTextToJson(const std::wstring& text);

#endif
