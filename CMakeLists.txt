cmake_minimum_required(VERSION 3.16)
project(ProcessAudioCapture VERSION 1.0.0 LANGUAGES CXX)

# C++17を使用
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Windows SDK設定
set(CMAKE_SYSTEM_VERSION 10.0.19041.0)

# 出力ディレクトリ
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)

# DLLターゲット
add_library(ProcessAudioCapture SHARED
    src/ProcessAudioCapture.cpp
    src/ProcessAudioCapture.h
    src/LoopbackCapture.cpp
    src/LoopbackCapture.h
)

# プリプロセッサ定義
target_compile_definitions(ProcessAudioCapture PRIVATE
    PROCESSAUDIOCAPTURE_EXPORTS
    UNICODE
    _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)

# インクルードディレクトリ
target_include_directories(ProcessAudioCapture PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/src
)

# リンクライブラリ
target_link_libraries(ProcessAudioCapture PRIVATE
    ole32
    mfplat
    propsys
    Psapi
    mmdevapi
)

# デバッグ情報を出力（デバッグビルド時）
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
    target_compile_options(ProcessAudioCapture PRIVATE /Zi)
    target_link_options(ProcessAudioCapture PRIVATE /DEBUG)
endif()

# Visual Studioでのソースグループ設定
source_group("Source Files" FILES
    src/ProcessAudioCapture.cpp
    src/LoopbackCapture.cpp
)
source_group("Header Files" FILES
    src/ProcessAudioCapture.h
    src/LoopbackCapture.h
)

# テスト用コンソールアプリケーション（オプション）
option(BUILD_TEST_APP "Build test application" ON)

if(BUILD_TEST_APP)
    add_executable(TestCapture
        test/test_capture.cpp
    )
    target_link_libraries(TestCapture PRIVATE
        ProcessAudioCapture
    )
    target_include_directories(TestCapture PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
    )
endif()

# インストール設定
install(TARGETS ProcessAudioCapture
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION lib
    ARCHIVE DESTINATION lib
)

install(FILES
    src/ProcessAudioCapture.h
    DESTINATION include
)

# パッケージング情報
set(CPACK_PACKAGE_NAME "ProcessAudioCapture")
set(CPACK_PACKAGE_VERSION ${PROJECT_VERSION})
set(CPACK_PACKAGE_DESCRIPTION_SUMMARY "Process-specific audio capture DLL for Windows")
include(CPack)
