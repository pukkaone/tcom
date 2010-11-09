# $Id: all.tcl,v 1.1 2002/03/16 04:53:17 cthuang Exp $
#
# This file contains a top-level script to run all of the tests.

if {[lsearch [namespace children] ::tcltest] == -1} {
    package require tcltest
    namespace import -force ::tcltest::*
}

set ::tcltest::testSingleFile false
set ::tcltest::testsDirectory [file dir [info script]]

foreach file [::tcltest::getMatchingFiles] {
    if {[catch {source $file} msg]} {
	puts stdout $msg
    }
}

::tcltest::cleanupTests 1
return
