# $Id: pkgIndex.tcl,v 1.16 2003/04/17 03:17:30 cthuang Exp $
package ifneeded tcom 3.9 \
[list load [file join $dir tcom.dll]]\n[list source [file join $dir tcom.tcl]]
