<?xml version="1.0"?>
<!-- $Id: refentry2html.xsl,v 1.16 2002/06/29 15:34:52 cthuang Exp $ -->
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
  <xsl:output method="html" encoding="UTF-8"
   doctype-public="-//W3C//DTD HTML 4.01 Transitional//EN"
   doctype-system="http://www.w3.org/TR/html4/loose.dtd"/>

  <xsl:template match="refentry">
    <html>
      <head>
        <title><xsl:value-of select="refnamediv/refname"/></title>
	<style type="text/css">
	  .command
	  { font-style: normal; font-weight: bold; }
	  .option
	  { font-style: normal; font-weight: bold; }
	  .parameter
	  { font-style: italic; font-weight: normal; }
	</style>
      </head>
      <body>
        <h2>Name</h2>
	<p><xsl:value-of select="refnamediv/refname"/> -- <xsl:value-of select="refnamediv/refpurpose"/></p>
	<xsl:apply-templates/>
      </body>
    </html>
  </xsl:template>

  <xsl:template match="docinfo"/>

  <xsl:template match="refmeta"/>

  <xsl:template match="refnamediv"/>

  <xsl:template match="refsynopsisdiv">
    <h2>Synopsis</h2>
    <xsl:apply-templates/>
  </xsl:template>

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
    <span class="option"><xsl:apply-templates/></span>
  </xsl:template>

  <xsl:template match="refsect1">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="refsect1/title">
    <h2><xsl:value-of select="text()"/></h2>
  </xsl:template>

  <xsl:template match="refsect2">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="refsect2/title">
    <h3><xsl:value-of select="text()"/></h3>
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

  <xsl:template match="parameter">
    <var><xsl:apply-templates/></var>
  </xsl:template>

  <xsl:template match="literal">
    <tt><xsl:apply-templates/></tt>
  </xsl:template>

  <xsl:template match="table">
    <blockquote>
    <xsl:apply-templates/>
    </blockquote>
  </xsl:template>

  <xsl:template match="table/title">
    <b><xsl:value-of select="."/></b><br/>
  </xsl:template>

  <xsl:template match="table/tgroup">
    <table border="1">
    <xsl:apply-templates/>
    </table>
  </xsl:template>

  <xsl:template match="thead">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="tbody">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template match="row">
    <tr><xsl:apply-templates/></tr>
  </xsl:template>

  <xsl:template match="thead/row/entry">
    <td><b><xsl:value-of select="."/></b></td>
  </xsl:template>

  <xsl:template match="tbody/row/entry">
    <td><xsl:value-of select="."/></td>
  </xsl:template>

  <xsl:template match="programlisting">
    <table bgcolor="#CCCCCC" width="100%"><tr><td><pre>
    <xsl:apply-templates/>
    </pre></td></tr></table>
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
