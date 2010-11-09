# $Id: events.tcl,v 1.2 2001/06/30 18:42:58 cthuang Exp $

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
