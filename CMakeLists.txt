cmake_minimum_required(VERSION 3.20)

project(VisualRecognitionAddIn
    VERSION 2.0.8
    LANGUAGES CXX
    DESCRIPTION "1C native add-in (VisualRecognition)"
)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

set(BOOST_ROOT "" CACHE PATH "Optional Boost root (directory containing boost/)")

# --- Target: Windows DLL for 1C ---
add_library(VisualRecognitionAddIn SHARED
    src/Component.cpp
    src/exports.cpp
    src/VisualAddIn.cpp
    src/GeminiPdfClient.cpp
    src/BoostJsonSingle.cpp
)

target_include_directories(VisualRecognitionAddIn
    PRIVATE
        "${CMAKE_CURRENT_SOURCE_DIR}/src"
        "${CMAKE_CURRENT_SOURCE_DIR}/third_party/1c-addin"
)
if(BOOST_ROOT)
    target_include_directories(VisualRecognitionAddIn PRIVATE "${BOOST_ROOT}")
endif()

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
else()
    # Linux HTTPS path for Gemini (Boost.Beast + OpenSSL).
    find_package(OpenSSL REQUIRED)
    find_package(Threads REQUIRED)
    target_link_libraries(VisualRecognitionAddIn PRIVATE OpenSSL::SSL OpenSSL::Crypto Threads::Threads)
    # Keep 1C naming convention on Linux: visualAddInLin32.so / visualAddInLin64.so.
    if(CMAKE_SIZEOF_VOID_P EQUAL 8)
        set(_VIS_ADDIN_OUTPUT_NAME visualAddInLin64)
    else()
        set(_VIS_ADDIN_OUTPUT_NAME visualAddInLin32)
    endif()
    set_target_properties(VisualRecognitionAddIn PROPERTIES
        OUTPUT_NAME "${_VIS_ADDIN_OUTPUT_NAME}"
        PREFIX ""
    )
endif()

if(MSVC)
    target_compile_options(VisualRecognitionAddIn PRIVATE /W3 /utf-8)
endif()

# DLL / import library: build/<preset>/bin/<Config>/, lib/<Config>/
set_target_properties(VisualRecognitionAddIn PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib/$<CONFIG>"
)
