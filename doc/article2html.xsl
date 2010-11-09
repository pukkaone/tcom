<?xml version="1.0"?>
<!-- $Id: article2html.xsl,v 1.11 2002/06/29 15:34:52 cthuang Exp $ -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html" encoding="UTF-8"
   doctype-public="-//W3C//DTD HTML 4.01 Transitional//EN"
   doctype-system="http://www.w3.org/TR/html4/loose.dtd"/>

  <xsl:template match="article">
    <html>
      <head>
        <title><xsl:value-of select="artheader/title"/></title>
	<style type="text/css">
	  .command
	  { font-style: normal; font-weight: bold; }
	  .option
	  { font-style: normal; font-weight: bold; }
	  .replaceable
	  { font-style: italic; font-weight: normal; }
	  .listing
	  { font-size: 9pt; }
	</style>
      </head>
      <body>
        <h1><xsl:value-of select="artheader/title"/></h1>
	<xsl:apply-templates/>
      </body>
    </html>
  </xsl:template>

  <xsl:template match="artheader"/>

  <xsl:template match="cmdsynopsis">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="arg">
    <var>
    <xsl:choose>
      <xsl:when test="@choice='plain'"></xsl:when>
      <xsl:otherwise>?</xsl:otherwise>
    </xsl:choose>
    <xsl:apply-templates/>
    <xsl:choose>
      <xsl:when test="@rep='repeat'"> ...</xsl:when>
    </xsl:choose>
    <xsl:choose>
      <xsl:when test="@choice='plain'"></xsl:when>
      <xsl:otherwise>?</xsl:otherwise>
    </xsl:choose>
    </var>
  </xsl:template>

  <xsl:template match="option">
    <code><xsl:apply-templates/></code>
  </xsl:template>

  <xsl:template match="sect1">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="sect1/title">
    <h2><xsl:value-of select="text()"/></h2>
  </xsl:template>

  <xsl:template match="command">
    <span class="command"><xsl:apply-templates/></span>
  </xsl:template>

  <xsl:template match="sbr">
    <br/>
  </xsl:template>

  <xsl:template match="variablelist">
    <dl>
      <xsl:apply-templates/>
    </dl>
  </xsl:template>

  <xsl:template match="varlistentry">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="varlistentry/term">
    <dt><xsl:apply-templates/></dt>
  </xsl:template>

  <xsl:template match="varlistentry/listitem">
    <dd><xsl:apply-templates/></dd>
  </xsl:template>

  <xsl:template match="para">
    <p><xsl:apply-templates/></p>
  </xsl:template>

  <xsl:template match="replaceable">
    <var><xsl:apply-templates/></var>
  </xsl:template>

  <xsl:template match="literal">
    <tt><xsl:apply-templates/></tt>
  </xsl:template>

  <xsl:template match="programlisting">
    <table bgcolor="#CCCCCC" width="100%"><tr><td><pre class="listing">
    <xsl:apply-templates/>
    </pre></td></tr></table>
  </xsl:template>

  <xsl:template match="screen">
    <table bgcolor="#FFFFCC" width="100%"><tr><td><pre>
    <xsl:apply-templates/>
    </pre></td></tr></table>
  </xsl:template>

  <xsl:template match="userinput">
    <kbd><xsl:apply-templates/></kbd>
  </xsl:template>

  <xsl:template match="mediaobject">
    <div><xsl:apply-templates/></div>
  </xsl:template>

  <xsl:template match="imageobject">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="imagedata">
    <img src="{@fileref}"/>
  </xsl:template>

  <xsl:template match="*">
    <font color="red">
      <xsl:text>&lt;</xsl:text>
      <xsl:value-of select="name(.)"/>
      <xsl:text>&gt;</xsl:text>
      <xsl:apply-templates/>
      <xsl:text>&lt;/</xsl:text>
      <xsl:value-of select="name(.)"/>
      <xsl:text>&gt;</xsl:text>
    </font>
  </xsl:template>

</xsl:stylesheet>
