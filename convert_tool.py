#coding: utf-8

# Imports
from lxml import etree
from StringIO import StringIO
import zipfile
import os, sys
import codecs
import shutil

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


# A decoder for Docx files
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


def extract_docx_images(docname, imgpath="." + os.sep):
    
    ## Get the zipfile
    z = zipfile.ZipFile(docname)

    ## Go through all files
    folderchecked = False
    
    for f in z.namelist():
        ## Extract all files with the right filename
        if f[11:16] == "image":
            ## Check if folder exists
            if not folderchecked:
                if not (os.path.exists(imgpath) and os.path.isdir(imgpath)):
                    os.mkdir(imgpath)
                    print " created folder " + imgpath
                    folderchecked = True
            
            ## Get filename
            filename = os.path.basename(f)
            print " Writing image " + filename
            
            ## Open source and target
            source = z.open(f)
            target = open(imgpath + os.sep + filename, "wb")
            
            ## Copy the data
            shutil.copyfileobj(source, target)
            
            ## CLose source and target
            source.close()
            target.close()
            
            
    

# A function that transforms docx files to txt files
def docx_indtt_file(filename, path="." + os.sep):
    
    print "Transforming " + filename
    decoded = decode_docx(filename, 'indtt-transform.xsl')
    
    if decoded:
        txtfilename = path + os.path.splitext(filename)[0] + '.txt'
        print " Writing to " + txtfilename
        try:
            outfile = codecs.open(txtfilename,'w','utf-16')
            outfile.write(decoded)
            outfile.close()
            print " Finished writing " + txtfilename
            
        except:
            print " Error!"
            
        imgpath = path + os.path.splitext(sys.argv[1])[0] + "_images"
        
        extract_docx_images(sys.argv[1], imgpath)
    
# Main
if len(sys.argv) == 2 and os.path.exists(sys.argv[1]):
    docx_indtt_file(sys.argv[1])
else:
    print "No existing file specified."
    
raw_input("Press ENTER to continue")