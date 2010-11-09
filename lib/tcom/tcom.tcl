# $Id: tcom.tcl,v 1.17 2003/04/02 22:46:51 cthuang Exp $

namespace eval ::tcom {
    # Tear down all event connections to the object.
    proc unbind {handle} {
	::tcom::bind $handle {}
    }

    # Look for the file in the directories in the package load path.
    # Return the full path of the file.
    proc search_auto_path {fileSpec} {
	global auto_path

	::foreach dir [set auto_path] {
	    set filePath [file join $dir $fileSpec]
	    if {[file exists $filePath]} {
		return [file nativename $filePath]
	    }
	}
	error "cannot find $fileSpec"
    }

    # Get full path to Tcl interpreter DLL.
    proc tclDllPath {} {
	set parts [file split [::info library]]
	set n [expr [llength $parts] - 3]
	set rootDir [eval file join [lrange $parts 0 $n]]
	set version [string map {. {}} [::info tclversion]]
	return [file nativename [file join $rootDir "bin" "tcl$version.dll"]]
    }

    # Insert registry entries for the class.
    proc registerClass {
	typeLibName typeLibId version className clsid inproc local
    } {
	set dllPath [search_auto_path "tcom/tcominproc.dll"]
	set exePath [search_auto_path "tcom/tcomlocal.exe"]
	if {[string first " " $exePath] > 0} {
	    # Must not have space character in local server path name.
	    set exePath [::tcom::shortPathName $exePath]
	}
	set verIndependentProgId "$typeLibName.$className"
	set progId "$verIndependentProgId.1"

	set key "HKEY_CLASSES_ROOT\\$progId"
	registry set $key "" "$className Class"
	registry set "$key\\CLSID" "" $clsid

	set key "HKEY_CLASSES_ROOT\\$verIndependentProgId"
	registry set $key "" "$className Class"
	registry set "$key\\CLSID" "" $clsid
	registry set "$key\\CurVer" "" $progId

	set key "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
	registry set $key "" "$className Class"
	registry set "$key\\ProgID" "" $progId
	registry set "$key\\VersionIndependentProgID" "" $verIndependentProgId
	registry set "$key\\TypeLib" "" $typeLibId
	registry set "$key\\Version" "" $version

	if {$inproc} {
	    set key "HKEY_CLASSES_ROOT\\CLSID\\$clsid\\InprocServer32"
	    registry set $key "" $dllPath
	}

	if {$local} {
	    set key "HKEY_CLASSES_ROOT\\CLSID\\$clsid\\LocalServer32"
	    registry set $key "" "$exePath $clsid"
	}

	set key "HKEY_CLASSES_ROOT\\CLSID\\$clsid\\tcom"
	registry set $key "Script" "package require $typeLibName"
	registry set $key "TclDLL" [tclDllPath]
    }

    # Remove registry entries for the class.
    proc unregisterClass {typeLibName className clsid} {
	set verIndependentProgId "$typeLibName.$className"
	set progId "$verIndependentProgId.1"

	registry delete "HKEY_CLASSES_ROOT\\$progId"
	registry delete "HKEY_CLASSES_ROOT\\$verIndependentProgId"
	registry delete "HKEY_CLASSES_ROOT\\CLSID\\$clsid"
    }

    # Register or unregister servers for classes defined in a type library.
    proc server {subCommand args} {
	package require registry

	set inproc 1
	set local 1

	set argc [llength $args]
	for {set i 0} {$i < $argc} {incr i} {
	    set endOfOptions 0
	    switch -- [lindex $args $i] {
		-inproc {
		    set inproc 1
		    set local 0
		}
		-local {
		    set inproc 0
		    set local 1
		}
		default {
		    set endOfOptions 1
		}
	    }
	    if {$endOfOptions} {
		break
	    }
	}

	if {$i >= $argc} {
	    error "wrong # args: usage: ::tcom::server register|unregister typeLibFile ?class ...?" 
	}

	set typeLibFile [lindex $args $i]
	incr i

	switch -- $subCommand {
	    register {
		::tcom::typelib register $typeLibFile
		set registerOpt 1
	    }
	    unregister {
		::tcom::typelib unregister $typeLibFile
		set registerOpt 0
	    }
	    default {
		error "bad option $option: must be register or unregsiter"
	    }
	}

	set typeLib [::tcom::typelib load $typeLibFile]
	set typeLibName [$typeLib name]
	set typeLibId "{[string toupper [$typeLib libid]]}"
	set typeLibVersion [$typeLib version]

	if {$i < $argc} {
	    set classes [lrange $args $i end]
	} else {
	    set classes [$typeLib class]
	}

	::foreach className $classes {
	    set classInfo [$typeLib class $className] 
	    set clsid "{[string toupper [lindex $classInfo 0]]}"
	    if {$registerOpt} {
		registerClass $typeLibName $typeLibId $typeLibVersion \
		    $className $clsid $inproc $local
	    } else {
		unregisterClass $typeLibName $className $clsid 
	    }
	}
    }
}
