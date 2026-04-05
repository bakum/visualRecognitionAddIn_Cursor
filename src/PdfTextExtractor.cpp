#include "PdfTextExtractor.h"

#if defined(VISUAL_ADDIN_HAVE_MUPDF)

#include <mupdf/fitz.h>

#ifdef _WIN32
#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#endif

#include <string>
#include <vector>

namespace {

#ifdef _WIN32
std::wstring Utf8ToWideZ(const char* utf8) {
    if (!utf8 || !utf8[0]) {
        return {};
    }
    const int n = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, utf8, -1, nullptr, 0);
    if (n <= 0) {
        return {};
    }
    std::wstring w(static_cast<size_t>(n - 1), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, utf8, -1, w.data(), n);
    return w;
}
#else
std::wstring Utf8ToWideZ(const char* utf8) {
    std::wstring out;
    if (!utf8) {
        return out;
    }
    while (*utf8) {
        const unsigned char c = static_cast<unsigned char>(*utf8++);
        if (c < 0x80) {
            out += static_cast<wchar_t>(c);
        } else {
            out += L'?';
        }
    }
    return out;
}
#endif

} // namespace

std::wstring ExtractPdfTextW(const std::vector<char>& pdf_bytes) {
    if (pdf_bytes.empty()) {
        return {};
    }

    fz_context* ctx = fz_new_context(nullptr, nullptr, FZ_STORE_UNLIMITED);
    if (!ctx) {
        return {};
    }

    fz_document* doc = nullptr;
    std::wstring result;
    fz_var(doc);

    fz_try(ctx) {
        fz_register_document_handlers(ctx);
        fz_stream* stm = fz_open_memory(
            ctx,
            reinterpret_cast<const unsigned char*>(pdf_bytes.data()),
            pdf_bytes.size());
        doc = fz_open_document_with_stream(ctx, "application/pdf", stm);
        fz_drop_stream(ctx, stm);

        const int page_count = fz_count_pages(ctx, doc);
        for (int i = 0; i < page_count; ++i) {
            fz_page* page = fz_load_page(ctx, doc, i);
            const fz_rect mediabox = fz_bound_page(ctx, page);
            fz_stext_page* stext = fz_new_stext_page(ctx, mediabox);
            fz_stext_options opts{};
            opts.flags = FZ_STEXT_PRESERVE_WHITESPACE | FZ_STEXT_PRESERVE_LIGATURES;
            fz_device* dev = fz_new_stext_device(ctx, stext, &opts);
            fz_run_page(ctx, page, dev, fz_identity, nullptr);
            fz_close_device(ctx, dev);
            fz_drop_device(ctx, dev);
            fz_drop_page(ctx, page);

            fz_buffer* buf = fz_new_buffer(ctx, 256);
            fz_output* out = fz_new_output_with_buffer(ctx, buf);
            fz_print_stext_page_as_text(ctx, out, stext);
            fz_drop_output(ctx, out);

            const char* txt = fz_string_from_buffer(ctx, buf);
            if (txt && txt[0] != '\0') {
                if (!result.empty()) {
                    result += L'\n';
                }
                result += Utf8ToWideZ(txt);
            }
            fz_drop_buffer(ctx, buf);
            fz_drop_stext_page(ctx, stext);
        }
    }
    fz_always(ctx) {
        fz_drop_document(ctx, doc);
    }
    fz_catch(ctx) {
        result.clear();
    }

    fz_drop_context(ctx);
    return result;
}

#elif defined(VISUAL_ADDIN_HAVE_PDFIUM)

#include <mutex>

#include <fpdf_text.h>
#include <fpdfview.h>

namespace {

std::once_flag g_pdfium_init;

void EnsurePdfiumInit() {
    std::call_once(g_pdfium_init, []() { FPDF_InitLibrary(); });
}

} // namespace

std::wstring ExtractPdfTextW(const std::vector<char>& pdf_bytes) {
    if (pdf_bytes.empty()) {
        return {};
    }

    EnsurePdfiumInit();

    FPDF_DOCUMENT doc = FPDF_LoadMemDocument64(
        pdf_bytes.data(),
        pdf_bytes.size(),
        nullptr);
    if (!doc) {
        return {};
    }

    std::wstring out;
    const int page_count = FPDF_GetPageCount(doc);
    for (int pi = 0; pi < page_count; ++pi) {
        FPDF_PAGE page = FPDF_LoadPage(doc, pi);
        if (!page) {
            continue;
        }
        FPDF_TEXTPAGE text_page = FPDFText_LoadPage(page);
        if (text_page) {
            const int n = FPDFText_CountChars(text_page);
            if (n > 0) {
                std::vector<unsigned short> buf(static_cast<size_t>(n) + 1u);
                const int written = FPDFText_GetText(text_page, 0, n, buf.data());
                if (written > 1) {
                    const size_t ulen = static_cast<size_t>(written > 0 ? written - 1 : 0);
                    out.append(reinterpret_cast<const wchar_t*>(buf.data()), ulen);
                }
            }
            FPDFText_ClosePage(text_page);
        }
        FPDF_ClosePage(page);
        if (pi + 1 < page_count) {
            out += L'\n';
        }
    }

    FPDF_CloseDocument(doc);
    return out;
}

#else

std::wstring ExtractPdfTextW(const std::vector<char>&) {
    return {};
}

#endif
