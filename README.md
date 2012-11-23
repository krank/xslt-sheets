This project contains XSLT experiments.

Current contents:
docx.xml (example docx xml)
docx-numbering.xml (example docx numbering xml, needed to see difference between bullet points and numbered lists)
indtt-transform.xsl (xslt transformer for docx/xml to indesign tagged text)


Notes on indtt-transform.xsl:
 - Its primary purpose is conversion of docx-documents downloaded från google docs. This means it will probably not work very well with a lot of "normal" docx documents.
 - No support for nested tables. Gdocs cannot export documents with nested tables to docx, which means support for nested tables is not a priority.
 - No support for highlighted text (text with colored background), since indesign tagged text lacks this cvapability
 - InDesign Tagged Text lacks image support. Images in the docx are replaced by a small text indicating their position.
 - Features are implemented one at a time, and there will probably always be some missing.
 - Some features, such as custom styles, will not be implemented because the purpose of the xslt is to get the document structure, not its layout, into InDesign.
 - The resulting InDesign Tagged Text document will contain a number of paragraph styles. These styles lack formatting.