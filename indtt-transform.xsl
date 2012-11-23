<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<xsl:output method="text"/>
<xsl:strip-space elements="*"/>

<!-- Numbering file to use -->
<xsl:variable name="numberingfile" select="'docx-numbering.xml'"/>
<xsl:variable name="scale" select="0.03"/>

<!-- Options -->
<xsl:variable name="use_justifications" select="1"/>



<!-- Basic document structure -->
<xsl:template match="/">
    <xsl:text>&lt;UNICODE-WIN&gt;&#13;</xsl:text>
    
    <!-- Set up paragraph styles -->
    <xsl:text>&lt;DefineParaStyle:NormalParagraphStyle&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Title&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Subtitle&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Heading1&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Heading2&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Heading3&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Heading4&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Heading5&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineParaStyle:Heading6&gt;&#13;</xsl:text>

    <!-- Set up character styles -->
    <xsl:text>&lt;DefineCharStyle:Italic=&lt;cTypeface:Italic&gt;&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineCharStyle:Bold=&lt;cTypeface:Bold&gt;&gt;&#13;</xsl:text>
    <xsl:text>&lt;DefineCharStyle:ItalicBold=&lt;cTypeface:Bold Italic&gt;&gt;&#13;</xsl:text>
    
    <xsl:apply-templates/>
</xsl:template>

<xsl:template match="w:tbl">
    <xsl:text>&lt;ParaStyle:NormalParagraphStyle&gt;</xsl:text>
    <xsl:text>&lt;TableStart:&gt;</xsl:text>
    
    <!-- Go through all rows -->
    <xsl:for-each select="w:tr">
        <xsl:text>&lt;RowStart:&gt;</xsl:text>
        
        <!-- Go through all the cells of each row -->
        <xsl:for-each select="w:tc">
            <xsl:text>&lt;CellStart:&gt;</xsl:text>
            
            <xsl:apply-templates select="w:p"/>
            
            <xsl:text>&lt;CellEnd:&gt;</xsl:text>
        </xsl:for-each>
        
        <xsl:text>&lt;RowEnd:&gt;</xsl:text>
    </xsl:for-each>
    
    
    <xsl:text>&lt;TableEnd:&gt;&#13;</xsl:text>
</xsl:template>

<!-- Match all paragraphs -->
<xsl:template match="w:p">
    <!-- Get the paragraph style -->
    <xsl:text>&lt;ParaStyle:</xsl:text>
    <xsl:if test="w:pPr/w:pStyle/@w:val != ''">
        <xsl:value-of select="w:pPr/w:pStyle/@w:val"/>
    </xsl:if>
    <xsl:if test="not(w:pPr/w:pStyle)">
        <xsl:text>NormalParagraphStyle</xsl:text>
    </xsl:if>
    
    <xsl:text>&gt;</xsl:text>
    
    <!-- Justifications -->
    <xsl:if test="$use_justifications = 1 and w:pPr/w:jc">
        <!-- Centered -->
        <xsl:choose>
            <xsl:when test="w:pPr/w:jc/@w:val = 'center'">
                <xsl:text>&lt;pTextAlignment:Center&gt;</xsl:text>
            </xsl:when>
            <xsl:when test="w:pPr/w:jc/@w:val = 'right'">
                <xsl:text>&lt;pTextAlignment:Right&gt;</xsl:text>
            </xsl:when>
            <xsl:when test="w:pPr/w:jc/@w:val = 'both'">
                <xsl:text>&lt;pTextAlignment:JustifyLeft&gt;</xsl:text>
            </xsl:when>-->
        </xsl:choose>        
        
    </xsl:if>
    
    
    
    <!-- Indents -->
    <xsl:if test="w:pPr/w:ind">
        <!-- Left indent -->
        <xsl:if test="w:pPr/w:ind/@w:left">
            <xsl:text>&lt;pLeftIndent:</xsl:text>
            <xsl:value-of select="round(w:pPr/w:ind/@w:left * $scale)"/>
            <xsl:text>&gt;</xsl:text>
        </xsl:if>
        <!-- First line -->
        <xsl:if test="w:pPr/w:ind/@w:hanging or w:pPr/w:ind/@w:firstLine">
            <xsl:text>&lt;pFirstLineIndent:</xsl:text>
            <xsl:choose>
                <xsl:when test="w:pPr/w:ind/@w:hanging and not(w:pPr/w:ind/@w:firstLine)">
                    <xsl:value-of select="round((w:pPr/w:ind/@w:hanging * -1) * $scale)"/>
                </xsl:when>
                <xsl:when test="not(w:pPr/w:ind/@w:hanging) and w:pPr/w:ind/@w:firstLine">
                    <xsl:value-of select="round(w:pPr/w:ind/@w:firstLine * $scale)"/>
                </xsl:when>
                <xsl:otherwise>
                    <xsl:value-of select="round(w:pPr/w:ind/@w:firstLine + (w:pPr/w:ind/@w:hanging * -1) * $scale)"/>
                </xsl:otherwise>
                
            </xsl:choose>
            <xsl:text>&gt;</xsl:text>
        </xsl:if>
    </xsl:if>
    
    
    <!-- Lists -->
    <xsl:if test="w:pPr/w:numPr">
        <xsl:call-template name="listtype">
            <xsl:with-param name="num">
                <xsl:value-of select="w:pPr/w:numPr/w:numId/@w:val"/>
            </xsl:with-param>
            <xsl:with-param name="level">
                <xsl:value-of select="w:pPr/w:numPr/w:ilvl/@w:val"/>
            </xsl:with-param>
        </xsl:call-template>
    </xsl:if>

    <!-- Get the contents to the paragraph text -->
    <xsl:apply-templates select="w:r"/>
    
    <!-- Add a CR after the paragraph, unless it's the last one -->
    <xsl:if test="not(position() = last())">
        <xsl:text>&#13;</xsl:text>
    </xsl:if>
    
</xsl:template>

<!-- Template for sub-r's -->
<xsl:template match="w:r">
    
    <!-- Character styles -->
    <xsl:if test="(w:rPr/w:b) or (w:rPr/w:i) or (w:rPr/w:u)">
        <xsl:text>&lt;CharStyle:</xsl:text>
    </xsl:if>
    
    <!-- Italic text -->
    <xsl:if test="w:rPr/w:i">
        <xsl:text>Italic</xsl:text>
    </xsl:if>
    
    <!-- Bold text -->    
    <xsl:if test="w:rPr/w:b">
        <xsl:text>Bold</xsl:text>
    </xsl:if>
    
    <xsl:if test="(w:rPr/w:b) or (w:rPr/w:i) or (w:rPr/w:u)">
        <xsl:text>&gt;</xsl:text>
    </xsl:if>
    
    <!-- Underlined text -->
    <xsl:if test="w:rPr/w:u">
        <xsl:text>&lt;cUnderline:1&gt;</xsl:text>
    </xsl:if>
    
    <!-- Get the text content -->
    <xsl:value-of select="."/>
    
    <xsl:if test="w:rPr/w:u">
        <xsl:text>&lt;cUnderline:&gt;</xsl:text>
    </xsl:if>
    
    <!-- Escape the formatting, if there was one. -->
    <xsl:if test="(w:rPr/w:b) or (w:rPr/w:i)">
        <xsl:text>&lt;CharStyle:&gt;</xsl:text>
    </xsl:if>
    
</xsl:template>


<!-- Get the list type of a list item -->
<xsl:template name="listtype">
    <xsl:param name="num" select="1"/>
    <xsl:param name="level" select="0"/>
    <xsl:param name="numberfile" select="'numbering.xml'"/>
    
    <xsl:text>&lt;bnListType:</xsl:text>
    <xsl:call-template name="abstractlist">
        <xsl:with-param name="level" select="$level"/>
        
        <xsl:with-param name="rnum">
            <xsl:value-of select="document($numberingfile)/w:numbering/w:num[@w:numId=$num]/w:abstractNumId/@w:val"/>
        </xsl:with-param>
    </xsl:call-template>
    <xsl:text>&gt;</xsl:text>
    
    
</xsl:template>

<xsl:template name="abstractlist">
    <xsl:param name="rnum" select="1"/>
    <xsl:param name="level" select="0"/>
    <xsl:call-template name="liststyleconvert">
        <xsl:with-param name="xstyle">
            <xsl:value-of select="document($numberingfile)/w:numbering/w:abstractNum[@w:abstractNumId=$rnum]/w:lvl[@w:ilvl=$level]/w:numFmt/@w:val"/>
        </xsl:with-param>
    </xsl:call-template>
</xsl:template>

<xsl:template name="liststyleconvert">
    <xsl:param name="xstyle" select="'bullet'"/>
    
    <xsl:choose>
        <xsl:when test="$xstyle='bullet'">
            <xsl:text>Bullet</xsl:text>
        </xsl:when>
        <xsl:otherwise>
            <xsl:text>Numbered</xsl:text>
        </xsl:otherwise>
    </xsl:choose>
    
</xsl:template>

<!--
Needs a' fixin:
 - Images
 - Tables w/ recursive paragraph management
 - Special entities(?)
Low priority:
 - Colored text
 - Super/sub
 - Footnotes

-->

</xsl:stylesheet>