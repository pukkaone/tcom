# $Id: pkgIndex.tcl,v 1.2 2002/03/30 18:49:10 cthuang Exp $
package ifneeded TclScript 1.0 \
[list load [file join $dir TclScript.dll]]\n[list source [file join $dir TclScript.itcl]]
