# $Id: xslt.tcl,v 1.2 2002/09/05 22:10:25 cthuang Exp $
#
# Run an XML document through an XSLT processor.

if {$argc != 3} {
   puts "usage: $argv0 inputFile xsltFile outputFile"
   exit 1
}

package require tcom

set domProgId "Msxml2.DOMDocument"

set source [::tcom::ref createobject $domProgId]
$source preserveWhiteSpace 1
$source validateOnParse 0
$source resolveExternals 0
set sourceUrl [lindex $argv 0]
if {![$source load $sourceUrl]} {
    set parseError [$source parseError]
    puts [format "%x" [$parseError errorCode]]
    puts [$parseError reason]
    puts [$parseError srcText]
    puts [$parseError url]
    exit 1
}

set xslt [::tcom::ref createobject $domProgId]
$xslt preserveWhiteSpace 1
$xslt validateOnParse 0
set xsltUrl [lindex $argv 1]
if {![$xslt load $xsltUrl]} {
    set parseError [$xslt parseError]
    puts [format "%x" [$parseError errorCode]]
    puts [$parseError reason]
    puts [$parseError srcText]
    puts [$parseError url]
    exit 1
}

regsub {<META http-equiv="Content-Type"[^>]*>} [$source transformNode $xslt] \
    {<META http-equiv="Content-Type" content="text/html; charset=UTF-8">} \
    resultHtml

set out [open [lindex $argv 2] "w"]
fconfigure $out -translation binary
puts -nonewline $out $resultHtml
close $out
