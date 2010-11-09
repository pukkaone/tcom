# $Id: events.tcl 5 2005-02-16 14:57:24Z cthuang $

package require tcom

proc sink {method args} {
    puts "event $method $args"
}

proc doUpdate {comment} {
    puts "invoked $comment"
    update
}

set application [::tcom::ref createobject "InternetExplorer.Application"]
::tcom::bind $application sink

$application Visible 1
doUpdate "Visible"
$application Quit
doUpdate "Quit"
