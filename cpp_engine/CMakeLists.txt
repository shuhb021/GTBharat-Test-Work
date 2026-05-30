cmake_minimum_required(VERSION 3.15)
project(far_engine LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# nlohmann/json — header-only, fetch automatically
include(FetchContent)
FetchContent_Declare(
    json
    GIT_REPOSITORY https://github.com/nlohmann/json.git
    GIT_TAG        v3.11.3
)
FetchContent_MakeAvailable(json)

# Build shared library (DLL on Windows)
add_library(far_engine SHARED
    src/variance_calc.cpp
    src/ratio_engine.cpp
    src/data_processor.cpp
)

target_include_directories(far_engine PRIVATE src)
target_link_libraries(far_engine PRIVATE nlohmann_json::nlohmann_json)

# Output to build/ directory
set_target_properties(far_engine PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/build"
    LIBRARY_OUTPUT_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/build"
)

# Install
install(TARGETS far_engine
    RUNTIME DESTINATION bin
    LIBRARY DESTINATION lib
)
