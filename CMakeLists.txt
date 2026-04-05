cmake_minimum_required(VERSION 3.20)

project(VisualRecognitionAddIn
    VERSION 2.0.2
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

# DLL / import library: build/<preset>/bin/<Config>/, lib/<Config>/
set_target_properties(VisualRecognitionAddIn PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib/$<CONFIG>"
)
