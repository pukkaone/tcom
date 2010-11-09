# $Id$
#
# This script unregisters the Tcl Active Scripting engine.

package require registry
package require tcom

    set typeLibFile "TclScript.tlb"
    ::tcom::server unregister $typeLibFile

    registry delete "HKEY_CLASSES_ROOT\\TclScript"
    registry delete "HKEY_CLASSES_ROOT\\.tcls"
    registry delete "HKEY_CLASSES_ROOT\\TclScriptFile"
