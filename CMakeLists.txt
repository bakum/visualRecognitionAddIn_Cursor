cmake_minimum_required(VERSION 3.20)

project(VisualRecognitionAddIn
    VERSION 2.0.13
    LANGUAGES CXX
    DESCRIPTION "1C native add-in (VisualRecognition)"
)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

option(VISUAL_ADDIN_BUILD_TESTS "Build unit tests" ON)

set(BOOST_ROOT "" CACHE PATH "Optional Boost root (directory containing boost/)")

set(_VIS_ADDIN_CORE_SOURCES
    src/Component.cpp
    src/VisualAddIn.cpp
    src/AnthropicMessagesClient.cpp
    src/GeminiPdfClient.cpp
    src/BoostJsonSingle.cpp
    src/AddInUtils.cpp
    src/ModelResolve.cpp
    src/ApiResponseParse.cpp
)

add_library(VisualRecognitionAddInCore STATIC ${_VIS_ADDIN_CORE_SOURCES})

if(NOT WIN32)
    set_target_properties(VisualRecognitionAddInCore PROPERTIES POSITION_INDEPENDENT_CODE ON)
endif()

target_include_directories(VisualRecognitionAddInCore
    PUBLIC
        "${CMAKE_CURRENT_SOURCE_DIR}/src"
        "${CMAKE_CURRENT_SOURCE_DIR}/third_party/1c-addin"
)

if(BOOST_ROOT)
    target_include_directories(VisualRecognitionAddInCore PUBLIC "${BOOST_ROOT}")
endif()

target_compile_definitions(VisualRecognitionAddInCore
    PUBLIC
        BOOST_ALL_NO_LIB
        "VISUAL_ADDIN_VERSION=\"${PROJECT_VERSION}\"")

if(WIN32)
    target_link_libraries(VisualRecognitionAddInCore PUBLIC winhttp)
    target_compile_definitions(VisualRecognitionAddInCore PUBLIC _WINDOWS)
else()
    find_package(OpenSSL REQUIRED)
    find_package(Threads REQUIRED)
    target_link_libraries(VisualRecognitionAddInCore PUBLIC OpenSSL::SSL OpenSSL::Crypto Threads::Threads)
endif()

if(MSVC)
    target_compile_options(VisualRecognitionAddInCore PRIVATE /W3 /utf-8)
endif()

# --- Target: Windows DLL for 1C ---
add_library(VisualRecognitionAddIn SHARED src/exports.cpp)
target_link_libraries(VisualRecognitionAddIn PRIVATE VisualRecognitionAddInCore)

if(WIN32)
    if(CMAKE_SIZEOF_VOID_P EQUAL 8)
        set(_VIS_ADDIN_OUTPUT_NAME visualAddInWin64)
    else()
        set(_VIS_ADDIN_OUTPUT_NAME visualAddInWin32)
    endif()
    set_target_properties(VisualRecognitionAddIn PROPERTIES
        OUTPUT_NAME "${_VIS_ADDIN_OUTPUT_NAME}"
        WINDOWS_EXPORT_ALL_SYMBOLS ON
    )
else()
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

set_target_properties(VisualRecognitionAddIn PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    ARCHIVE_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/lib/$<CONFIG>"
)

# --- Unit tests (Catch2) ---
if(VISUAL_ADDIN_BUILD_TESTS)
    include(FetchContent)
    FetchContent_Declare(
        Catch2
        GIT_REPOSITORY https://github.com/catchorg/Catch2.git
        GIT_TAG v3.5.4
    )
    FetchContent_MakeAvailable(Catch2)

    add_executable(VisualRecognitionAddInTests
        tests/test_addin_utils.cpp
        tests/test_model_resolve.cpp
        tests/test_api_response_parse.cpp
        tests/test_primary_document_normalize.cpp
        tests/test_visual_addin_methods.cpp
    )

    target_link_libraries(VisualRecognitionAddInTests
        PRIVATE
            VisualRecognitionAddInCore
            Catch2::Catch2WithMain
    )

    if(MSVC)
        target_compile_options(VisualRecognitionAddInTests PRIVATE /W3 /utf-8)
    endif()

    set_target_properties(VisualRecognitionAddInTests PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY "${CMAKE_BINARY_DIR}/bin/$<CONFIG>"
    )

    include(CTest)
    include(Catch)
    catch_discover_tests(VisualRecognitionAddInTests)
endif()
