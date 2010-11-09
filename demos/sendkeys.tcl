# $Id: sendkeys.tcl 5 2005-02-16 14:57:24Z cthuang $
#
# This example demonstrates how to send keys to Windows applications.
# It requires Windows Script Host 2.0 installed on the system.

package require tcom

set wshShell [::tcom::ref createobject "WScript.Shell"]
set taskId [$wshShell Run "notepad.exe"]
$wshShell AppActivate $taskId
after 500
$wshShell SendKeys "The quick brown fox jumped\n"
$wshShell SendKeys "{TAB}over the lazy dog."
