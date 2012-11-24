This project contains XSLT experiments.

Current contents:
docx.xml (example docx xml)
docx-numbering.xml (example docx numbering xml, needed to see difference between bullet points and numbered lists)
indtt-transform.xsl (xslt transformer for docx/xml to indesign tagged text)


Notes on indtt-transform.xsl:
 - Its primary purpose is conversion of docx-documents downloaded from google docs. This means it will probably not work very well with a lot of "normal" docx documents.
 - Features are implemented one at a time, and there will probably always be some missing.
 - Some features will not be implemented because the purpose of the xslt is to get the document structure, not its layout, into InDesign.
 - The resulting InDesign Tagged Text document will contain a number of paragraph styles. These styles lack formatting.
 - No support for nested tables. Gdocs cannot export documents with nested tables to docx, which means support for nested tables is not a priority.
 
InDesign Tagged Text is a more limited format than Google Docs or Docx. This means:
 - No image support. Images in the docx are replaced by a small text indicating their position.
 - No footnotes within tables. Footnotes within tables will look like this: <?>.
 - No highlighted text (text with colored background).