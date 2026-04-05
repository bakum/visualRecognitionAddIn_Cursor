# Downloads and extracts prebuilt PDFium (bblanchon/pdfium-binaries) for Windows.
# After extraction, ${pdfium_extract_dir} contains PDFiumConfig.cmake, include/, lib/, bin/.

set(PDFIUM_TAG "chromium/7763" CACHE STRING "pdfium-binaries release tag path (e.g. chromium/7763)")

if(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(_PDFIUM_ARCHIVE "pdfium-win-x64.tgz")
else()
    set(_PDFIUM_ARCHIVE "pdfium-win-x86.tgz")
endif()

set(_PDFIUM_URL "https://github.com/bblanchon/pdfium-binaries/releases/download/${PDFIUM_TAG}/${_PDFIUM_ARCHIVE}")
string(REPLACE ".tgz" "" _PDFIUM_DIRNAME "${_PDFIUM_ARCHIVE}")
set(pdfium_extract_dir "${CMAKE_BINARY_DIR}/_deps/${_PDFIUM_DIRNAME}")

if(EXISTS "${pdfium_extract_dir}/PDFiumConfig.cmake")
    return()
endif()

message(STATUS "Fetching PDFium (${_PDFIUM_ARCHIVE})...")
file(MAKE_DIRECTORY "${pdfium_extract_dir}")

set(_PDFIUM_ARCHIVE_FILE "${CMAKE_BINARY_DIR}/_deps/${_PDFIUM_ARCHIVE}")
file(DOWNLOAD "${_PDFIUM_URL}" "${_PDFIUM_ARCHIVE_FILE}"
    TLS_VERIFY ON
    STATUS _dl_status
)
list(GET _dl_status 0 _dl_code)
if(NOT _dl_code EQUAL 0)
    list(GET _dl_status 1 _dl_msg)
    message(FATAL_ERROR "PDFium download failed (${_dl_code}): ${_dl_msg}")
endif()

execute_process(
    COMMAND "${CMAKE_COMMAND}" -E tar xfz "${_PDFIUM_ARCHIVE_FILE}"
    WORKING_DIRECTORY "${pdfium_extract_dir}"
    RESULT_VARIABLE _extract_rc
)
if(NOT _extract_rc EQUAL 0)
    message(FATAL_ERROR "Extracting PDFium archive failed (${_extract_rc})")
endif()

if(NOT EXISTS "${pdfium_extract_dir}/PDFiumConfig.cmake")
    message(FATAL_ERROR "PDFiumConfig.cmake not found after extract (broken archive?)")
endif()

message(STATUS "PDFium ready at ${pdfium_extract_dir}")
