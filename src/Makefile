# $Id: Makefile,v 1.12 2003/07/24 22:46:35 cthuang Exp $

debug:
	tclsh &&|
set libDir [file join [info library] "../tcom"]
file copy -force ../lib/tcom/pkgIndex.tcl $libDir
file copy -force ../lib/tcom/tcom.tcl $libDir
file copy -force Debug/tcom.dll $libDir
file copy -force dllserver_Debug/tcominproc.dll $libDir
file copy -force exeserver_Debug/tcomlocal.exe $libDir
set libDir [file join [info library] "../TclScript"]
file copy -force TclScript_Debug/TclScript.dll $libDir
file copy -force TclScript_Debug/TclScript.tlb $libDir
|

release:
	tclsh &&|
set libDir [file join [info library] "../tcom"]
file copy -force ../lib/tcom/pkgIndex.tcl $libDir
file copy -force ../lib/tcom/tcom.tcl $libDir
file copy -force Release/tcom.dll $libDir
file copy -force dllserver_Release/tcominproc.dll $libDir
file copy -force exeserver_Release/tcomlocal.exe $libDir
set libDir [file join [info library] "../TclScript"]
file copy -force TclScript_Release/TclScript.dll $libDir
file copy -force TclScript_Release/TclScript.tlb $libDir
|
