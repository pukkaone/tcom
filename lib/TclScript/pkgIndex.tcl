# $Id: pkgIndex.tcl 5 2005-02-16 14:57:24Z cthuang $
package ifneeded TclScript 1.0 \
[list load [file join $dir TclScript.dll]]\n[list source [file join $dir TclScript.itcl]]
