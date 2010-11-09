# $Id: pkgIndex.tcl 5 2005-02-16 14:57:24Z cthuang $
package ifneeded tcom 3.9 \
[list load [file join $dir tcom.dll]]\n[list source [file join $dir tcom.tcl]]
