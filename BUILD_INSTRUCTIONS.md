# Prerequisites

* ActiveTcl installed in the C:\Tcl directory
* CMake 2.8 or better
* Visual C++ 2008 Express Edition


# Building

Starting in the top distribution directory (where this BUILD_INSTRUCTIONS.md
file is located), run the commands:

    mkdir build
    cd build
    cmake -G "Visual Studio 9 2008" ..

In Visual C++ 2008, open the solution `tcom.sln`. Build the `ALL_BUILD` project.
