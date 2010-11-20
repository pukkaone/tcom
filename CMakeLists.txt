# $Id$

cmake_minimum_required(VERSION 2.8)

project(tcom)
include(cmake/configure_files.cmake)

add_subdirectory(src)

find_package(Git)
if(GIT_FOUND)
    # Get the package version from the git tag.
    execute_process(
            COMMAND ${GIT_EXECUTABLE} describe --abbrev=0
            OUTPUT_VARIABLE PACKAGE_VERSION
            OUTPUT_STRIP_TRAILING_WHITESPACE)

    # Get the build number by counting commits from the first commit in the
    # repository which is tagged "build".
    execute_process(
            COMMAND ${GIT_EXECUTABLE} describe --match build
            OUTPUT_VARIABLE DESCRIBE_BUILD
            OUTPUT_STRIP_TRAILING_WHITESPACE)
    string(REGEX MATCH "[0-9]+" BUILD_NUMBER ${DESCRIBE_BUILD})
else()
    set(PACKAGE_VERSION "0.0")
    set(PACKAGE_BUILD_NUMBER 0)
endif()

string(REGEX MATCH "[0-9]+" PACKAGE_MAJOR_VERSION ${PACKAGE_VERSION})
string(REPLACE
        "${PACKAGE_MAJOR_VERSION}."
        ""
        PACKAGE_MINOR_VERSION
        ${PACKAGE_VERSION})

configure_files(
        ${CMAKE_SOURCE_DIR}/template/include ${CMAKE_BINARY_DIR}/include)
