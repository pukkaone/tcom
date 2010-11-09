# $Id: sendkeys.tcl,v 1.3 2001/06/30 18:42:58 cthuang Exp $
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
