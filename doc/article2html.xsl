<?xml version="1.0"?>
<!-- $Id$ -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html" encoding="UTF-8"
   doctype-public="-//W3C//DTD HTML 4.01 Transitional//EN"
   doctype-system="http://www.w3.org/TR/html4/loose.dtd"/>

  <xsl:template match="article">
    <html>
      <head>
        <title><xsl:value-of select="title"/></title>
	<style type="text/css">
	  .command
	  { font-style: normal; font-weight: bold; }
	  .option
	  { font-style: normal; font-weight: bold; }
	  .programlisting
	  { background-color: #E8E8E8; font-size: 9pt; }
	  .screen
	  { background-color: #FFFFCC; }
	</style>
      </head>
      <body>
	<xsl:apply-templates/>
      </body>
    </html>
  </xsl:template>

  <xsl:template match="article/title">
    <h1><xsl:value-of select="text()"/></h1>
  </xsl:template>

  <xsl:template match="articleinfo"/>

  <xsl:template match="cmdsynopsis">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="arg">
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
    <pre class="programlisting">
    <xsl:apply-templates/>
    </pre>
  </xsl:template>

  <xsl:template match="screen">
    <pre class="screen">
    <xsl:apply-templates/>
    </pre>
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
