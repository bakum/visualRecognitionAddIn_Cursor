#include "InvoiceTextParser.h"

#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <Windows.h>

#include <boost/json.hpp>

#include <algorithm>
#include <cctype>
#include <cwctype>
#include <regex>
#include <string>
#include <vector>

namespace {

std::string WideToUtf8(const std::wstring& w) {
    if (w.empty()) {
        return {};
    }
    const int n = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), static_cast<int>(w.size()), nullptr, 0,
                                      nullptr, nullptr);
    if (n <= 0) {
        return {};
    }
    std::string out(static_cast<size_t>(n), '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), static_cast<int>(w.size()), out.data(), n, nullptr,
                        nullptr);
    return out;
}

void TrimInPlace(std::wstring& s) {
    while (!s.empty() && std::iswspace(static_cast<wint_t>(s.front()))) {
        s.erase(s.begin());
    }
    while (!s.empty() && std::iswspace(static_cast<wint_t>(s.back()))) {
        s.pop_back();
    }
}

std::vector<std::wstring> SplitLines(const std::wstring& text) {
    std::vector<std::wstring> lines;
    std::wstring cur;
    for (wchar_t ch : text) {
        if (ch == L'\r') {
            continue;
        }
        if (ch == L'\n') {
            TrimInPlace(cur);
            lines.push_back(std::move(cur));
            cur.clear();
        } else {
            cur += ch;
        }
    }
    TrimInPlace(cur);
    lines.push_back(std::move(cur));
    return lines;
}

std::wstring JoinNonEmpty(const std::vector<std::wstring>& parts, size_t from, size_t to) {
    std::wstring r;
    for (size_t i = from; i < to && i < parts.size(); ++i) {
        if (parts[i].empty()) {
            continue;
        }
        if (!r.empty()) {
            r += L' ';
        }
        r += parts[i];
    }
    return r;
}

bool LineLooksLikeTableHeader(const std::wstring& line) {
    const bool has_name = line.find(L"Товари") != std::wstring::npos
        || line.find(L"товари") != std::wstring::npos
        || line.find(L"Назва") != std::wstring::npos
        || line.find(L"найменування") != std::wstring::npos;
    const bool has_qty = line.find(L"Кіл-сть") != std::wstring::npos
        || line.find(L"Кількість") != std::wstring::npos;
    if (has_name && has_qty) {
        return true;
    }
    // «№ Товари (роботи, послуги) Кіл-сть Од. Ціна Сума»
    static const std::wregex hdr_num_goods(L"№\\s+.*Товар", std::regex::icase);
    return std::regex_search(line, hdr_num_goods) && line.find(L"Ціна") != std::wstring::npos;
}

bool LineLooksLikeFooter(const std::wstring& line) {
    static const std::wregex re(
        L"(Всього|Усього|Разом|До\\s+сплати|"
        L"Сторінка\\s+\\d|Стор\\.\\s*\\d|--\\s*\\d+\\s+of\\s+\\d+\\s*--)",
        std::regex::icase);
    return std::regex_search(line, re);
}

std::vector<std::wstring> SplitCells(const std::wstring& line) {
    if (line.find(L'\t') != std::wstring::npos) {
        std::vector<std::wstring> cells;
        std::wstring cur;
        for (wchar_t ch : line) {
            if (ch == L'\t') {
                TrimInPlace(cur);
                cells.push_back(std::move(cur));
                cur.clear();
            } else {
                cur += ch;
            }
        }
        TrimInPlace(cur);
        cells.push_back(std::move(cur));
        return cells;
    }
    static const std::wregex split_re(L"\\s{2,}");
    std::vector<std::wstring> cells(
        std::wsregex_token_iterator(line.begin(), line.end(), split_re, -1),
        std::wsregex_token_iterator());
    for (auto& c : cells) {
        TrimInPlace(c);
    }
    cells.erase(std::remove_if(cells.begin(), cells.end(),
                               [](const std::wstring& x) { return x.empty(); }),
                cells.end());
    return cells;
}

static const std::wregex kNumToken(
    L"^[+-]?(?:\\d{1,3}(?:\\s\\d{3})+|\\d+)(?:[.,]\\d+)?$");

bool IsNumericToken(const std::wstring& t) {
    std::wstring s = t;
    for (auto& ch : s) {
        if (ch == L' ') {
            ch = 0;
        }
    }
    s.erase(std::remove(s.begin(), s.end(), wchar_t(0)), s.end());
    return std::regex_match(s, kNumToken);
}

bool TryParseAmountTailLine(const std::wstring& line, std::wstring& qty, std::wstring& unit,
                            std::wstring& price, std::wstring& amount) {
    static const std::wregex re(
        L"^(\\d+)\\s+(грн\\.?|₴|USD|EUR)\\s+([\\d\\s]+[.,]\\d{2})\\s+([\\d\\s]+[.,]\\d{2})\\s*$",
        std::regex::icase);
    std::wsmatch m;
    if (!std::regex_match(line, m, re)) {
        return false;
    }
    qty = m[1].str();
    unit = m[2].str();
    price = m[3].str();
    amount = m[4].str();
    TrimInPlace(qty);
    TrimInPlace(unit);
    TrimInPlace(price);
    TrimInPlace(amount);
    return true;
}

boost::json::object LineItemJsonFromParts(const std::wstring& name, const std::wstring& qty,
                                          const std::wstring& unit, const std::wstring& price,
                                          const std::wstring& amount) {
    boost::json::object o;
    o["name"] = WideToUtf8(name);
    o["quantity"] = WideToUtf8(qty);
    o["unit"] = WideToUtf8(unit);
    o["price"] = WideToUtf8(price);
    o["amount"] = WideToUtf8(amount);
    return o;
}

std::wstring StripLeadingRowNumber(const std::wstring& s) {
    static const std::wregex strip(L"^\\d+\\s+");
    std::wstring out = std::regex_replace(s, strip, L"");
    TrimInPlace(out);
    return out;
}

boost::json::object TryInferLineItemJson(const std::vector<std::wstring>& cells) {
    if (cells.size() < 3) {
        return {};
    }
    std::vector<size_t> num_idx;
    for (size_t i = 0; i < cells.size(); ++i) {
        if (IsNumericToken(cells[i])) {
            num_idx.push_back(i);
        }
    }
    if (num_idx.size() >= 2) {
        const size_t last = num_idx.back();
        const size_t prev = num_idx[num_idx.size() - 2];
        std::wstring amount = cells[last];
        std::wstring price = cells[prev];
        std::wstring quantity;
        size_t name_end = prev;
        if (num_idx.size() >= 3) {
            const size_t qix = num_idx[num_idx.size() - 3];
            quantity = cells[qix];
            name_end = qix;
        }
        std::wstring name = JoinNonEmpty(cells, 0, name_end);
        TrimInPlace(name);
        if (name.empty()) {
            name = JoinNonEmpty(cells, 0, prev);
        }
        boost::json::object o;
        o["name"] = WideToUtf8(name);
        o["quantity"] = WideToUtf8(quantity);
        o["price"] = WideToUtf8(price);
        o["amount"] = WideToUtf8(amount);
        return o;
    }
    std::wstring raw = JoinNonEmpty(cells, 0, cells.size());
    boost::json::object o;
    o["name"] = WideToUtf8(raw);
    return o;
}

std::wstring CaptureFirst(const std::wstring& text, const std::wregex& re) {
    std::wsmatch m;
    if (std::regex_search(text, m, re) && m.size() > 1) {
        return m[1].str();
    }
    return {};
}

std::wstring BlockAfterLabel(const std::vector<std::wstring>& lines, const std::wstring& label) {
    for (size_t i = 0; i < lines.size(); ++i) {
        if (lines[i].find(label) == std::wstring::npos) {
            continue;
        }
        std::wstring block;
        for (size_t j = i; j < lines.size() && j < i + 8; ++j) {
            std::wstring ln = lines[j];
            if (j == i) {
                const size_t p = ln.find(label);
                if (p != std::wstring::npos) {
                    ln = ln.substr(p + label.size());
                    TrimInPlace(ln);
                    if (!ln.empty() && ln[0] == L':') {
                        ln.erase(ln.begin());
                        TrimInPlace(ln);
                    }
                }
            }
            if (ln.empty()) {
                if (!block.empty()) {
                    break;
                }
                continue;
            }
            static const std::wregex next_hdr(
                L"^(Покупець|Постачальник|Замовник|Виконавець|"
                L"Вантажоодержувач|Вантажовідправник|Відправник|Отримувач|"
                L"Банк|МФО|IBAN|Договір|Рахунок|Акт|"
                L"код за ДРФО|ЄДРПОУ|ІПН|Ідентифікаційний|п/р|"
                L"Адреса|Тел\\.|Телефон|Ел\\.\\s*пошта|e-mail)",
                std::regex::icase);
            if (j > i && std::regex_search(ln, next_hdr)) {
                break;
            }
            if (!block.empty()) {
                block += L' ';
            }
            block += ln;
            if (block.length() > 400) {
                break;
            }
        }
        TrimInPlace(block);
        return block;
    }
    return {};
}

} // namespace

std::string ParseInvoiceTextToJson(const std::wstring& text) {
    const auto lines = SplitLines(text);
    std::wstring flat = text;
    std::replace(flat.begin(), flat.end(), L'\r', L' ');
    std::replace(flat.begin(), flat.end(), L'\n', L' ');

    // ЄДРПОУ (юрособа), ІПН / ДРФО / ід. код — у JSON полі inn для сумісності з існуючою схемою.
    static const std::wregex re_edrpou(L"(?:ЄДРПОУ|EDRPOU)\\s*[:\\s]*(\\d{8})\\b", std::regex::icase);
    static const std::wregex re_ipn(L"(?:ІПН|IPN)\\s*[:\\s]*(\\d{10})\\b", std::regex::icase);
    static const std::wregex re_drho(L"код за ДРФО\\s+(\\d{8,12})\\b", std::regex::icase);
    static const std::wregex re_id_code(L"Ідентифікаційний\\s+код\\s*:?\\s*(\\d{8,12})\\b",
                                        std::regex::icase);
    static const std::wregex re_mfo(L"(?:МФО|MFO)\\s*[:\\s]*(\\d{6})\\b", std::regex::icase);
    static const std::wregex re_iban_ua(L"(?:п/р|р/р|IBAN)\\s*[:\\s]*(UA\\d{27})\\b",
                                         std::regex::icase);
    static const std::wregex re_contract(L"Договір\\s*:\\s*([^\\n\\r]+?)(?=\\s{2,}|$)",
                                         std::regex::icase);
    static const std::wregex re_contract_date(
        L"від\\s+(\\d{1,2}[./]\\d{1,2}[./]\\d{2,4})\\b", std::regex::icase);
    static const std::wregex re_invoice_payment(
        L"Рахунок на оплату\\s*№?\\s*(\\d+)", std::regex::icase);
    static const std::wregex re_invoice_rf(L"Рахунок-фактура\\s*№?\\s*(\\d+)", std::regex::icase);
    static const std::wregex re_bank_inline_ua(
        L"у банку\\s+([^,\\n\\r]{3,120})", std::regex::icase);

    std::wstring inn = CaptureFirst(flat, re_edrpou);
    if (inn.empty()) {
        inn = CaptureFirst(flat, re_ipn);
    }
    if (inn.empty()) {
        inn = CaptureFirst(flat, re_drho);
    }
    if (inn.empty()) {
        inn = CaptureFirst(flat, re_id_code);
    }
    const std::wstring kpp; // у первинках України поле КПП відсутнє
    const std::wstring bik = CaptureFirst(flat, re_mfo);
    const std::wstring rs = CaptureFirst(flat, re_iban_ua);
    const std::wstring ks; // кореспондентський рахунок банку — за потреби доповнити окремим шаблоном

    std::wstring supplier = BlockAfterLabel(lines, L"Постачальник");
    if (supplier.empty()) {
        supplier = BlockAfterLabel(lines, L"Виконавець");
    }
    if (supplier.empty()) {
        supplier = BlockAfterLabel(lines, L"Продавець");
    }
    std::wstring buyer = BlockAfterLabel(lines, L"Покупець");
    if (buyer.empty()) {
        buyer = BlockAfterLabel(lines, L"Замовник");
    }
    if (buyer.empty()) {
        buyer = BlockAfterLabel(lines, L"Вантажоодержувач");
    }

    std::wstring bank_line = BlockAfterLabel(lines, L"Банк отримувача");
    if (bank_line.empty()) {
        bank_line = BlockAfterLabel(lines, L"Банк");
    }
    if (bank_line.empty()) {
        bank_line = CaptureFirst(flat, re_bank_inline_ua);
    }

    std::wstring contract = CaptureFirst(flat, re_contract);
    TrimInPlace(contract);
    std::wstring contract_date = CaptureFirst(flat, re_contract_date);

    std::wstring invoice_number = CaptureFirst(flat, re_invoice_payment);
    if (invoice_number.empty()) {
        invoice_number = CaptureFirst(flat, re_invoice_rf);
    }

    boost::json::array items_json;
    bool in_table = false;
    /// Рядок позиції з ведучим номером «N …» (багаторядкова назва + окремий рядок сум у грн, або одна
    /// таблична лінія з кількома колонками).
    std::wstring pending_accum;

    auto append_item = [&](const boost::json::object& chunk) { items_json.push_back(chunk); };

    auto flush_pending_line_item_row = [&]() {
        if (pending_accum.empty()) {
            return;
        }
        const auto cells = SplitCells(pending_accum);
        const boost::json::object inferred = TryInferLineItemJson(cells);
        if (!inferred.empty()) {
            append_item(inferred);
        } else {
            boost::json::object one;
            one["name"] = WideToUtf8(StripLeadingRowNumber(pending_accum));
            append_item(one);
        }
        pending_accum.clear();
    };

    static const std::wregex row_start(L"^(\\d+)\\s+(.+)$");

    for (size_t i = 0; i < lines.size(); ++i) {
        if (!in_table) {
            if (LineLooksLikeTableHeader(lines[i])) {
                in_table = true;
            }
            continue;
        }
        if (lines[i].empty()) {
            continue;
        }
        if (LineLooksLikeFooter(lines[i])) {
            flush_pending_line_item_row();
            break;
        }
        if (LineLooksLikeTableHeader(lines[i])) {
            continue;
        }

        std::wstring q, u, pr, am;
        if (TryParseAmountTailLine(lines[i], q, u, pr, am)) {
            if (!pending_accum.empty()) {
                append_item(
                    LineItemJsonFromParts(StripLeadingRowNumber(pending_accum), q, u, pr, am));
                pending_accum.clear();
            }
            continue;
        }

        std::wsmatch rm;
        if (std::regex_match(lines[i], rm, row_start)) {
            flush_pending_line_item_row();
            pending_accum = lines[i];
            TrimInPlace(pending_accum);
            continue;
        }

        if (!pending_accum.empty()) {
            pending_accum += L' ';
            pending_accum += lines[i];
        } else {
            const auto cells = SplitCells(lines[i]);
            if (cells.size() >= 2) {
                const boost::json::object row = TryInferLineItemJson(cells);
                if (!row.empty()) {
                    append_item(row);
                }
            }
        }
    }
    flush_pending_line_item_row();

    boost::json::object counterparty;
    counterparty["supplier"] = WideToUtf8(supplier);
    counterparty["buyer"] = WideToUtf8(buyer);
    counterparty["inn"] = WideToUtf8(inn);
    counterparty["kpp"] = WideToUtf8(kpp);

    boost::json::object contract_obj;
    contract_obj["numberAndDetails"] = WideToUtf8(contract);
    contract_obj["date"] = WideToUtf8(contract_date);

    boost::json::object bank;
    bank["bankName"] = WideToUtf8(bank_line);
    bank["bik"] = WideToUtf8(bik);
    bank["correspondentAccount"] = WideToUtf8(ks);
    bank["settlementAccount"] = WideToUtf8(rs);

    boost::json::object root;
    root["counterparty"] = counterparty;
    root["contract"] = contract_obj;
    root["invoiceNumber"] = WideToUtf8(invoice_number);
    root["bankDetails"] = bank;
    root["lineItems"] = items_json;
    root["rawText"] = WideToUtf8(text);

    return std::string(boost::json::serialize(root));
}
