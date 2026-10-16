cmake_minimum_required(VERSION 3.20)
project(Tools28Render LANGUAGES C CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
if(NOT CMAKE_BUILD_TYPE)
  set(CMAKE_BUILD_TYPE Release)
endif()

include(FetchContent)

# Vulkan のヘッダー（GPU を動かす仕組みの定義）。Vulkan 本体は GPU ドライバーに入っているので読み込みは volk が行う
FetchContent_Declare(vulkan_headers
  GIT_REPOSITORY https://github.com/KhronosGroup/Vulkan-Headers.git
  GIT_TAG v1.4.309
  GIT_SHALLOW TRUE)
FetchContent_MakeAvailable(vulkan_headers)

# ウィンドウ作成（GLFW）。Linux では入っているものを使い、Windows ではソースからビルドする
option(T28_SYSTEM_GLFW "Use the GLFW installed on the system" OFF)
if(T28_SYSTEM_GLFW)
  find_package(glfw3 3.3 REQUIRED)
else()
  set(GLFW_BUILD_DOCS OFF CACHE BOOL "" FORCE)
  set(GLFW_BUILD_TESTS OFF CACHE BOOL "" FORCE)
  set(GLFW_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
  set(GLFW_INSTALL OFF CACHE BOOL "" FORCE)
  FetchContent_Declare(glfw
    GIT_REPOSITORY https://github.com/glfw/glfw.git
    GIT_TAG 3.4
    GIT_SHALLOW TRUE)
  FetchContent_MakeAvailable(glfw)
endif()

# シェーダー: glslc があればビルド時に変換し、無ければ変換済みのもの（shaders/generated）を使う
#   pathtrace_hw: RT コアで当たり判定（Vulkan 1.2 + ray query）
#   pathtrace_sw: BVH を自分でたどる予備の方式（Vulkan 1.0 で動く）
#   temporal / atrous / composite: ノイズ除去と仕上げ
#   overlay:      操作の案内を重ねる
find_program(GLSLC glslc)
set(SHADER_DIR ${CMAKE_CURRENT_SOURCE_DIR}/shaders)
set(SHADER_COMMON ${SHADER_DIR}/common_head.glsl ${SHADER_DIR}/common_body.glsl)
set(SHADER_DEPS "")
if(GLSLC)
  set(SHADER_OUT_DIR ${CMAKE_CURRENT_BINARY_DIR}/shaders)
  file(MAKE_DIRECTORY ${SHADER_OUT_DIR})
  foreach(spec pathtrace_hw:vulkan1.2 pathtrace_sw:vulkan1.0 temporal:vulkan1.0 atrous:vulkan1.0 composite:vulkan1.0 overlay:vulkan1.0)
    string(REPLACE ":" ";" pair ${spec})
    list(GET pair 0 name)
    list(GET pair 1 env)
    add_custom_command(
      OUTPUT ${SHADER_OUT_DIR}/${name}.spv.inc
      COMMAND ${GLSLC} --target-env=${env} -O -mfmt=c -I ${SHADER_DIR} ${SHADER_DIR}/${name}.comp -o ${SHADER_OUT_DIR}/${name}.spv.inc
      COMMAND ${CMAKE_COMMAND} -E copy ${SHADER_OUT_DIR}/${name}.spv.inc ${SHADER_DIR}/generated/${name}.spv.inc
      DEPENDS ${SHADER_DIR}/${name}.comp ${SHADER_COMMON}
      COMMENT "Compiling ${name}.comp")
    list(APPEND SHADER_DEPS ${SHADER_OUT_DIR}/${name}.spv.inc)
  endforeach()
  set(SHADER_INCLUDE_DIR ${SHADER_OUT_DIR})
else()
  set(SHADER_INCLUDE_DIR ${SHADER_DIR}/generated)
endif()

add_executable(Tools28Render
  src/main.cpp
  src/bvh.cpp
  src/help_overlay.cpp
  src/log.cpp
  src/i18n.cpp
  src/scene.cpp
  src/vk_context.cpp
  src/renderer.cpp
  src/offscreen.cpp
  src/window_app.cpp
  third_party/volk/volk.c
  ${SHADER_DEPS})

target_include_directories(Tools28Render PRIVATE
  src third_party/volk third_party/stb ${SHADER_INCLUDE_DIR})
target_link_libraries(Tools28Render PRIVATE Vulkan::Headers glfw)

if(WIN32)
  # 画面のないコンソールを出さない GUI アプリとして作る。実行に余分な DLL が要らないよう静的にリンクする
  target_compile_definitions(Tools28Render PRIVATE VK_USE_PLATFORM_WIN32_KHR UNICODE _UNICODE NOMINMAX)
  if(MINGW)
    target_link_options(Tools28Render PRIVATE -municode -mwindows -static -static-libgcc -static-libstdc++)
  else()
    set_target_properties(Tools28Render PROPERTIES WIN32_EXECUTABLE ON
      MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
  endif()
  target_link_libraries(Tools28Render PRIVATE shell32)
else()
  target_link_libraries(Tools28Render PRIVATE ${CMAKE_DL_LIBS} pthread)
endif()
