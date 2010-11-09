# $Id: register.tcl,v 1.3 2002/03/20 23:52:35 cthuang Exp $
#
# This script registers the Tcl Active Scripting engine.

package require registry
package require tcom

    set typeLibFile "TclScript.tlb"
    ::tcom::server register -inproc $typeLibFile

    set typeLib [::tcom::typelib load $typeLibFile]
    set classInfo [$typeLib class "Engine"]
    set clsid "{[string toupper [lindex $classInfo 0]]}"
    set progId "TclScript"

    set key "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
    registry set "$key\\ProgID" "" $progId
    registry set "$key\\OLEScript"

    set key "HKEY_CLASSES_ROOT\\CLSID\\$clsid\\Implemented Categories"
    registry set "$key\\{F0B7A1A1-9847-11CF-8F20-00805F2CD064}"
    registry set "$key\\{F0B7A1A2-9847-11CF-8F20-00805F2CD064}"

    set key "HKEY_CLASSES_ROOT\\$progId"
    registry set $key "" "Tcl Script Language"
    registry set "$key\\CLSID" "" $clsid
    registry set "$key\\OLEScript"

    set key "HKEY_CLASSES_ROOT\\.tcls"
    registry set $key "" "TclScriptFile"

    set key "HKEY_CLASSES_ROOT\\TclScriptFile"
    registry set $key "" "Tcl Script File"
    registry set "$key\\ScriptEngine" "" $progId
