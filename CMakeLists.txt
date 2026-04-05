cmake_minimum_required(VERSION 3.20)

project(VisualRecognitionAddIn
    VERSION 1.0.0
    LANGUAGES CXX
    DESCRIPTION "1C native add-in (VisualRecognition)"
)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

set(BOOST_ROOT "E:/boost/boost_1_84_0" CACHE PATH "Root of unpacked Boost (directory containing boost/)")

# --- Target: Windows DLL for 1C ---
add_library(VisualRecognitionAddIn SHARED
    src/Component.cpp
    src/exports.cpp
    src/VisualAddIn.cpp
    src/PdfTextExtractor.cpp
    src/InvoiceTextParser.cpp
    src/GeminiPdfClient.cpp
    src/BoostJsonSingle.cpp
)

target_include_directories(VisualRecognitionAddIn
    PRIVATE
        "${CMAKE_CURRENT_SOURCE_DIR}/src"
        "${CMAKE_CURRENT_SOURCE_DIR}/third_party/1c-addin"
        "${BOOST_ROOT}"
)

target_compile_definitions(VisualRecognitionAddIn PRIVATE
    BOOST_ALL_NO_LIB
    "VISUAL_ADDIN_VERSION=\"${PROJECT_VERSION}\"")

if(WIN32)
    target_link_libraries(VisualRecognitionAddIn PRIVATE winhttp)
    # SDK headers + __stdcall / WCHAR_T=wchar_t (see third_party/1c-addin/types.h)
    target_compile_definitions(VisualRecognitionAddIn PRIVATE _WINDOWS)
    # DLL name follows the bitness of the build (separate x86 / x64 presets)
    if(CMAKE_SIZEOF_VOID_P EQUAL 8)
        set(_VIS_ADDIN_OUTPUT_NAME visualAddInWin64)
    else()
        set(_VIS_ADDIN_OUTPUT_NAME visualAddInWin32)
    endif()
    # Export GetClassNames / GetClassObject / DestroyObject / SetPlatformCapabilities without a .def file
    set_target_properties(VisualRecognitionAddIn PROPERTIES
        OUTPUT_NAME "${_VIS_ADDIN_OUTPUT_NAME}"
        WINDOWS_EXPORT_ALL_SYMBOLS ON
    )
endif()

if(MSVC)
    target_compile_options(VisualRecognitionAddIn PRIVATE /W3 /utf-8)
endif()

# PDF: MuPDF (vcpkg unofficial-libmupdf) if found; otherwise fetch PDFium — build still works without vcpkg.
if(WIN32)
    option(VISUAL_ADDIN_WITH_PDF "Enable PDF text extraction (MuPDF if found, else PDFium)" ON)
    if(VISUAL_ADDIN_WITH_PDF)
        find_package(unofficial-libmupdf CONFIG QUIET)
        if(TARGET unofficial::libmupdf::libmupdf)
            target_compile_definitions(VisualRecognitionAddIn PRIVATE VISUAL_ADDIN_HAVE_MUPDF)
            target_link_libraries(VisualRecognitionAddIn PRIVATE unofficial::libmupdf::libmupdf)
            message(STATUS "PDF text engine: MuPDF (unofficial-libmupdf)")
        else()
            option(VISUAL_ADDIN_EMBED_PDFIUM
                "Embed pdfium.dll inside the add-in (single deployable DLL; extracted under %LOCALAPPDATA%)"
                ON)
            list(APPEND CMAKE_MODULE_PATH "${CMAKE_CURRENT_SOURCE_DIR}/cmake")
            include(AcquirePDFium)
            set(_PDFIUM_ROOT "${pdfium_extract_dir}")
            set(_PDFIUM_INCLUDE "${_PDFIUM_ROOT}/include")
            set(_PDFIUM_DLL "${_PDFIUM_ROOT}/bin/pdfium.dll")
            set(_PDFIUM_LIB "${_PDFIUM_ROOT}/lib/pdfium.dll.lib")
            if(NOT EXISTS "${_PDFIUM_INCLUDE}/fpdfview.h" OR NOT EXISTS "${_PDFIUM_LIB}")
                message(FATAL_ERROR "PDFium layout unexpected under ${_PDFIUM_ROOT}")
            endif()
            add_library(pdfium SHARED IMPORTED GLOBAL)
            set_target_properties(pdfium PROPERTIES
                IMPORTED_LOCATION "${_PDFIUM_DLL}"
                IMPORTED_IMPLIB "${_PDFIUM_LIB}"
                INTERFACE_INCLUDE_DIRECTORIES "${_PDFIUM_INCLUDE};${_PDFIUM_INCLUDE}/cpp"
            )
            target_compile_definitions(VisualRecognitionAddIn PRIVATE VISUAL_ADDIN_HAVE_PDFIUM)
            target_link_libraries(VisualRecognitionAddIn PRIVATE pdfium)
            if(VISUAL_ADDIN_EMBED_PDFIUM)
                # rc.exe: path to pdfium.dll at build time (embedded in our DLL as RCDATA)
                file(TO_CMAKE_PATH "${_PDFIUM_DLL}" _PDFIUM_RC_PATH)
                file(WRITE "${CMAKE_CURRENT_BINARY_DIR}/pdfium_embed.rc" "101 RCDATA \"${_PDFIUM_RC_PATH}\"\n")
                target_sources(VisualRecognitionAddIn PRIVATE
                    "${CMAKE_CURRENT_BINARY_DIR}/pdfium_embed.rc"
                    src/pdfium_delayload.cpp
                )
                target_compile_definitions(VisualRecognitionAddIn PRIVATE VISUAL_ADDIN_EMBED_PDFIUM)
                target_link_options(VisualRecognitionAddIn PRIVATE /DELAYLOAD:pdfium.dll)
                target_link_libraries(VisualRecognitionAddIn PRIVATE delayimp.lib)
                message(STATUS "PDF text engine: PDFium embedded (single deployable DLL + delay-load)")
            else()
                add_custom_command(TARGET VisualRecognitionAddIn POST_BUILD
                    COMMAND "${CMAKE_COMMAND}" -E copy_if_different
                        "${_PDFIUM_DLL}"
                        "$<TARGET_FILE_DIR:VisualRecognitionAddIn>"
                    COMMENT "Copy pdfium.dll next to add-in"
                )
                message(STATUS "PDF text engine: PDFium (pdfium.dll beside add-in)")
            endif()
        endif()
    endif()
endif()

# DLL / import library: build/<preset>/bin/<Config>/, lib/<Config>/
set_target_properties(VisualRecognitionAddIn PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib/$<CONFIG>"
)
