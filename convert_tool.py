# Imports
from lxml import etree
from StringIO import StringIO
import zipfile
import os
import codecs

# An URL resolver that finds files in a specified zip file
class ZipFileResolver(etree.Resolver):
    
    ## Set zipfile at creation
    def __init__(self, zfile):
        self.zfile = zfile

    ## Resolve
    def resolve(self, url, pubid, context):
        ## If the file exists, return its contents
        if url in self.zfile.namelist():
            contents = self.zfile.read(url)
        else:
            ## If not, return an empty "error" element
            contents = "<error/>"
        return self.resolve_string(contents, context)


def decode_docx(docname, xsltname):
    
    ## If file exists
    if os.path.exists(docname):
    
        ## Open it as a zip
        z = zipfile.ZipFile(docname)
        
        ## If the file seems to be in the docx format
        if 'word/document.xml' in z.namelist():
            
            ## And the specified xslt file exists
            if os.path.exists(xsltname):
            
                ## Create a new parser, and give it a zipfile resolver.
                docxParser = etree.XMLParser()
                docxParser.resolvers.add(ZipFileResolver(z))

                xsltfile = open(xsltname, "r")
                xmlfile = z.open('word/document.xml')
                
                parsed_xml = etree.parse(xmlfile, docxParser)
                parsed_xslt = etree.parse(xsltfile, docxParser)
                
                xsltfile.close()
                xmlfile.close()
                
                transformer = etree.XSLT(parsed_xslt)
                return unicode(transformer(parsed_xml))
                
                
            else:
                print "XSLT file not found"
        else:
            print "Not a word document."
    else:
        print "File doesn't exist"

decoded = decode_docx('exempel.docx', 'indtt-transform.xsl')

if decoded:
    outfile = codecs.open('exempel.txt','w','utf-16')
    outfile.write(decoded)
    outfile.close()