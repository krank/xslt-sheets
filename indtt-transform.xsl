<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
 xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<xsl:output method="text"/>
<xsl:strip-space elements="*"/>

<!-- Files to use -->
<xsl:variable name="numberingfile" select="'docx-numbering.xml'"/>
<xsl:variable name="footnotefile" select="'docx-footnotes.xml'"/>

<!-- Options -->
<xsl:variable name="use_justifications" select="1"/>
<xsl:variable name="scale" select="0.03"/>

<!-- Keys-->
<xsl:key name="styles" match="w:p/w:pPr/w:pStyle" use="@w:val"/>

<!-- Basic document structure -->
<xsl:template match="/">
    <xsl:text>&lt;UNICODE-WIN&gt;&#13;</xsl:text>
    
    <!-- Set up paragraph styles -->
    <xsl:text>&lt;DefineParaStyle:NormalParagraphStyle&gt;&#13;</xsl:text>

    <xsl:for-each select="//w:pStyle[generate-id() = generate-id(key('styles',@w:val)[1])]">
        <xsl:text>&lt;DefineParaStyle:</xsl:text>
        <xsl:value-of select="@w:val"/>
        <xsl:text>&gt;&#13;</xsl:text>
    </xsl:for-each>
    
    <xsl:apply-templates/>
</xsl:template>

<!-- Match all tables -->
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

<!-- Template for images -->

<xsl:template match="w:drawing">
    <xsl:text>&#91;IMAGE:</xsl:text>
    
    <xsl:if test="wp:inline"/>
    <xsl:choose>
        <xsl:when test="wp:inline">
            <xsl:text>INLINE:</xsl:text>
            <xsl:value-of select="wp:inline/wp:docPr/@name"/>
        </xsl:when>
        <xsl:when test="wp:anchor">
            <xsl:text>ANCHORED:</xsl:text>
            <xsl:value-of select="wp:anchor/wp:docPr/@name"/>
        </xsl:when>
    </xsl:choose>
    
    <xsl:text>&#93;</xsl:text>
</xsl:template>

<!-- Template for sub-r's -->
<xsl:template match="w:r">

    <!-- Begin bold/italic -->
    <xsl:choose>
        <xsl:when test="w:rPr/w:b and not(w:rPr/w:i)">
            <xsl:text>&lt;cTypeface:Bold&gt;</xsl:text>
        </xsl:when>
        <xsl:when test="not(w:rPr/w:b) and w:rPr/w:i">
            <xsl:text>&lt;cTypeface:Italic&gt;</xsl:text>
        </xsl:when>
        <xsl:when test="w:rPr/w:b and w:rPr/w:i">
            <xsl:text>&lt;cTypeface:Bold Italic&gt;</xsl:text>
        </xsl:when>
    </xsl:choose>
    
    <!-- Begin Underlined text -->
    <xsl:if test="w:rPr/w:u">
        <xsl:text>&lt;cUnderline:1&gt;</xsl:text>
    </xsl:if>
    
    <!-- Begin Strikethrough text -->
    <xsl:if test="w:rPr/w:strike">
        <xsl:text>&lt;cStrikethru:1&gt;</xsl:text>
    </xsl:if>
    
    <!-- Begin sub/superscript -->
    <xsl:if test="w:rPr/w:vertAlign and not(w:footnoteReference)">
        <xsl:text>&lt;cPosition:</xsl:text>
        <xsl:choose>
            <xsl:when test="w:rPr/w:vertAlign/@w:val = 'superscript'">
                <xsl:text>Superscript</xsl:text>
            </xsl:when>
            <xsl:when test = "w:rPr/w:vertAlign/@w:val = 'subscript'">
                <xsl:text>Subscript</xsl:text>
            </xsl:when>
        </xsl:choose>
        <xsl:text>&gt;</xsl:text>
    </xsl:if>
    
    <!-- Check if r contains a drawing or not -->
    <xsl:choose>
        <xsl:when test="w:drawing">
            <xsl:apply-templates match="w:drawing"/>
        </xsl:when>
        <xsl:otherwise>
            <!-- If it does not, get text content-->
            <xsl:for-each select="descendant::*">
                <xsl:choose>
                    <xsl:when test="name(.) = 'w:t'">
                        <xsl:value-of select="."/>
                    </xsl:when>
                    <xsl:when test="name(.) = 'w:tab'">
                        <xsl:text>&#09;</xsl:text>
                    </xsl:when>
                    <xsl:when test="name(.) = 'w:br' and @w:type='textWrapping'">
                        <xsl:text>&#13;</xsl:text>
                    </xsl:when>
                </xsl:choose>
            </xsl:for-each>
        </xsl:otherwise>
    </xsl:choose>

    <!-- End sub/superscript -->
    <xsl:if test="w:rPr/w:vertAlign and not(w:footnoteReference)">
        <xsl:text>&lt;cPosition:&gt;</xsl:text>
    </xsl:if>
    
    <!-- End strikethrough text -->
    <xsl:if test="w:rPr/w:strike">
        <xsl:text>&lt;cStrikethru:&gt;</xsl:text>
    </xsl:if>
    
    <!-- End underlined text -->
    <xsl:if test="w:rPr/w:u">
        <xsl:text>&lt;cUnderline:&gt;</xsl:text>
    </xsl:if>

    <!-- End bold/italic -->
    <xsl:if test="(w:rPr/w:b) or (w:rPr/w:i)">
        <xsl:text>&lt;cTypeface:&gt;</xsl:text>
    </xsl:if>
    
    <!-- Footnotes -->
    <xsl:if test="w:footnoteReference">
        <xsl:call-template name="footnote">
            <xsl:with-param name="id" select="w:footnoteReference/@w:id"/>
        </xsl:call-template>
    </xsl:if>
    
    <!-- Page break -->
    <xsl:if test="w:br">
            <xsl:if test="not(w:br/@w:type='textWrapping')">
                <xsl:text>&lt;cNextXChars:Column&gt;</xsl:text>
            </xsl:if>
    </xsl:if>
            
</xsl:template>


<!-- Footnotes -->
<xsl:template name="footnote">
    <xsl:param name="id" select="0"/>
    
    <xsl:text>&lt;FootnoteStart:&gt;</xsl:text>
    <xsl:value-of select="document($footnotefile)/w:footnotes/w:footnote[@w:id = $id]"/>
    <xsl:text>&lt;FootnoteEnd:&gt;</xsl:text>
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

</xsl:stylesheet>