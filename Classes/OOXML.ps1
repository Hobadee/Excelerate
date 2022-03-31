class OOXML : IDisposable {

#https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e
#https://docs.fileformat.com/spreadsheet/xlsx/
#https://docs.fileformat.com/word-processing/docx/
#https://docs.fileformat.com/presentation/pptx/


<#

The "Excel" class should largely be refactored into this.  Excel should then extend this if anything.
Read the Schema docs linked above before designing.



# Flow
1.  Check that file is a zip file
2.  Unzip file into temp dir
3.  Check for and load [Content_Types].xml from root dir.  If it doesn't exist, it isn't an OOXML.
4.  Verify OOXML by checking top of [Content_Types].xml: 
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
5.  Load all "PartName" into array of "PartName" objects
6.  Search for "PartName" where "ContentType" in:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" (Excel)
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" (Word)
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" (PowerPoint)
7.  If we find that we are Excel, continue to load, otherwise exception
8.  Load "PartName" where "ContentType" = application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml
    (This should be "xl/workbook.xml")
    This is our "Excel" object
9.  In our Excel object, map "externalReferences" using "r:id"="id" to xl/_rels/workbook.xml.rels
10. Check "Target" of all references, split path and load externalLinks/_rels/<Target>.xml.rel


#>

    <##
     # Implement Dispose to fullfill the IDisposable interface
    #>
    Dispose(){
        #base.Dispose  # This is supposed to be called, but errors out
    }


}
