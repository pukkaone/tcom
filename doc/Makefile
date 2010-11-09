# $Id: Makefile,v 1.6 2002/04/17 22:07:57 cthuang Exp $

all: tcom.n.html server.html

tcom.n.html: tcom.n.xml xslt.tcl refentry2html.xsl
	tclsh xslt.tcl tcom.n.xml refentry2html.xsl $@

server.html: server.xml xslt.tcl article2html.xsl
	tclsh xslt.tcl server.xml article2html.xsl $@
