# Microsoft Developer Studio Project File - Name="tcom" - Package Owner=<4>
# Microsoft Developer Studio Generated Build File, Format Version 6.00
# ** DO NOT EDIT **

# TARGTYPE "Win32 (x86) Dynamic-Link Library" 0x0102

CFG=tcom - Win32 No DCOM Debug
!MESSAGE This is not a valid makefile. To build this project using NMAKE,
!MESSAGE use the Export Makefile command and run
!MESSAGE 
!MESSAGE NMAKE /f "tcom.mak".
!MESSAGE 
!MESSAGE You can specify a configuration when running NMAKE
!MESSAGE by defining the macro CFG on the command line. For example:
!MESSAGE 
!MESSAGE NMAKE /f "tcom.mak" CFG="tcom - Win32 No DCOM Debug"
!MESSAGE 
!MESSAGE Possible choices for configuration are:
!MESSAGE 
!MESSAGE "tcom - Win32 Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "tcom - Win32 Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "tcom - Win32 No DCOM Release" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE "tcom - Win32 No DCOM Debug" (based on "Win32 (x86) Dynamic-Link Library")
!MESSAGE 

# Begin Project
# PROP AllowPerConfigDependencies 0
# PROP Scc_ProjName ""
# PROP Scc_LocalPath ""
CPP=cl.exe
MTL=midl.exe
RSC=rc.exe

!IF  "$(CFG)" == "tcom - Win32 Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "Release"
# PROP BASE Intermediate_Dir "Release"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "Release"
# PROP Intermediate_Dir "Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /D "WIN32" /D "NDEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /YX /FD /c
# ADD CPP /nologo /MD /W3 /GX /Zi /Od /I "\tcl\include" /D "NDEBUG" /D "_WIN32_DCOM" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /D "TCL_THREADS" /D "USE_TCL_STUBS" /D "USE_NON_CONST" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386
# ADD LINK32 rpcrt4.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /libpath:"\tcl\lib"

!ELSEIF  "$(CFG)" == "tcom - Win32 Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "Debug"
# PROP BASE Intermediate_Dir "Debug"
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "Debug"
# PROP Intermediate_Dir "Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /D "WIN32" /D "_DEBUG" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /YX /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /I "\tcl\include" /D "_DEBUG" /D "_WIN32_DCOM" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /D "TCL_THREADS" /D "USE_TCL_STUBS" /D "USE_NON_CONST" /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept
# ADD LINK32 rpcrt4.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /libpath:"\tcl\lib"

!ELSEIF  "$(CFG)" == "tcom - Win32 No DCOM Release"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 0
# PROP BASE Output_Dir "tcom___Win32_No_DCOM_Release"
# PROP BASE Intermediate_Dir "tcom___Win32_No_DCOM_Release"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 0
# PROP Output_Dir "No_DCOM_Release"
# PROP Intermediate_Dir "No_DCOM_Release"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MT /W3 /GX /O2 /I "e:\tcl\include" /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /D "_WIN32_DCOM" /YX /FD /c
# ADD CPP /nologo /MD /W3 /GX /Zi /O2 /I "\tcl\include" /D "NDEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /D "TCL_THREADS" /D "USE_TCL_STUBS" /D "USE_NON_CONST" /YX /FD /c
# ADD BASE MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "NDEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "NDEBUG"
# ADD RSC /l 0x409 /d "NDEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 tk80.lib tcl80.lib rpcrt4.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /machine:I386 /libpath:"e:\tcl\lib"
# ADD LINK32 rpcrt4.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /libpath:"\tcl\lib"

!ELSEIF  "$(CFG)" == "tcom - Win32 No DCOM Debug"

# PROP BASE Use_MFC 0
# PROP BASE Use_Debug_Libraries 1
# PROP BASE Output_Dir "tcom___Win32_No_DCOM_Debug"
# PROP BASE Intermediate_Dir "tcom___Win32_No_DCOM_Debug"
# PROP BASE Ignore_Export_Lib 0
# PROP BASE Target_Dir ""
# PROP Use_MFC 0
# PROP Use_Debug_Libraries 1
# PROP Output_Dir "No_DCOM_Debug"
# PROP Intermediate_Dir "No_DCOM_Debug"
# PROP Ignore_Export_Lib 0
# PROP Target_Dir ""
# ADD BASE CPP /nologo /MTd /W3 /Gm /GX /ZI /Od /I "c:\tcl\include" /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /D "_WIN32_DCOM" /YX /FD /GZ /c
# ADD CPP /nologo /MDd /W3 /Gm /GX /Zi /Od /I "\tcl\include" /D "_DEBUG" /D "WIN32" /D "_WINDOWS" /D "_MBCS" /D "_USRDLL" /D "TCOM_EXPORTS" /D "TCL_THREADS" /D "USE_TCL_STUBS" /D "USE_NON_CONST" /YX /FD /GZ /c
# ADD BASE MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD MTL /nologo /D "_DEBUG" /mktyplib203 /win32
# ADD BASE RSC /l 0x409 /d "_DEBUG"
# ADD RSC /l 0x409 /d "_DEBUG"
BSC32=bscmake.exe
# ADD BASE BSC32 /nologo
# ADD BSC32 /nologo
LINK32=link.exe
# ADD BASE LINK32 tk80.lib tcl80.lib rpcrt4.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /pdbtype:sept /libpath:"c:\tcl\lib"
# ADD LINK32 rpcrt4.lib kernel32.lib user32.lib gdi32.lib winspool.lib comdlg32.lib advapi32.lib shell32.lib ole32.lib oleaut32.lib uuid.lib odbc32.lib odbccp32.lib /nologo /dll /debug /machine:I386 /libpath:"\tcl\lib"

!ENDIF 

# Begin Target

# Name "tcom - Win32 Release"
# Name "tcom - Win32 Debug"
# Name "tcom - Win32 No DCOM Release"
# Name "tcom - Win32 No DCOM Debug"
# Begin Group "Source Files"

# PROP Default_Filter "cpp;c;cxx;rc;def;r;odl;idl;hpj;bat"
# Begin Source File

SOURCE=.\Arguments.cpp
# End Source File
# Begin Source File

SOURCE=.\bindCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\ComModule.cpp
# End Source File
# Begin Source File

SOURCE=.\ComObject.cpp
# End Source File
# Begin Source File

SOURCE=.\ComObjectFactory.cpp
# End Source File
# Begin Source File

SOURCE=.\configureCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\Extension.cpp
# End Source File
# Begin Source File

SOURCE=.\foreachCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\HandleSupport.cpp
# End Source File
# Begin Source File

SOURCE=.\importCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\infoCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\InterfaceAdapter.cpp
# End Source File
# Begin Source File

SOURCE=.\InterfaceAdapterVtbl.cpp
# End Source File
# Begin Source File

SOURCE=.\main.cpp
# End Source File
# Begin Source File

SOURCE=.\naCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\nullCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\objectCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\refCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\Reference.cpp
# End Source File
# Begin Source File

SOURCE=.\RegistryKey.cpp
# End Source File
# Begin Source File

SOURCE=.\shortPathNameCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\SupportErrorInfo.cpp
# End Source File
# Begin Source File

SOURCE=.\TclObject.cpp
# End Source File
# Begin Source File

SOURCE=.\tcomVersion.rc
# End Source File
# Begin Source File

SOURCE=.\TypeInfo.cpp
# End Source File
# Begin Source File

SOURCE=.\TypeLib.cpp
# End Source File
# Begin Source File

SOURCE=.\typelibCmd.cpp
# End Source File
# Begin Source File

SOURCE=.\Uuid.cpp
# End Source File
# Begin Source File

SOURCE=.\variantCmd.cpp
# End Source File
# End Group
# Begin Group "Header Files"

# PROP Default_Filter "h;hpp;hxx;hm;inl"
# Begin Source File

SOURCE=.\Arguments.h
# End Source File
# Begin Source File

SOURCE=.\ComModule.h
# End Source File
# Begin Source File

SOURCE=.\ComObject.h
# End Source File
# Begin Source File

SOURCE=.\ComObjectFactory.h
# End Source File
# Begin Source File

SOURCE=.\Extension.h
# End Source File
# Begin Source File

SOURCE=.\HandleSupport.h
# End Source File
# Begin Source File

SOURCE=.\HashTable.h
# End Source File
# Begin Source File

SOURCE=.\InterfaceAdapter.h
# End Source File
# Begin Source File

SOURCE=.\mutex.h
# End Source File
# Begin Source File

SOURCE=.\Reference.h
# End Source File
# Begin Source File

SOURCE=.\RegistryKey.h
# End Source File
# Begin Source File

SOURCE=.\Singleton.h
# End Source File
# Begin Source File

SOURCE=.\SupportErrorInfo.h
# End Source File
# Begin Source File

SOURCE=.\TclObject.h
# End Source File
# Begin Source File

SOURCE=.\tclRunTime.h
# End Source File
# Begin Source File

SOURCE=.\tcomApi.h
# End Source File
# Begin Source File

SOURCE=.\ThreadLocalStorage.h
# End Source File
# Begin Source File

SOURCE=.\TypeInfo.h
# End Source File
# Begin Source File

SOURCE=.\TypeLib.h
# End Source File
# Begin Source File

SOURCE=.\Uuid.h
# End Source File
# End Group
# Begin Group "Resource Files"

# PROP Default_Filter "ico;cur;bmp;dlg;rc2;rct;bin;rgs;gif;jpg;jpeg;jpe"
# End Group
# End Target
# End Project
