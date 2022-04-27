class OOXML : IDisposable {

#https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e
#https://docs.fileformat.com/spreadsheet/xlsx/
#https://docs.fileformat.com/word-processing/docx/
#https://docs.fileformat.com/presentation/pptx/


<#

The "Excel" class should largely be refactored into this.  Excel should then extend this if anything.
Read the Schema docs linked above before designing.



# Flow
1.  [X] Check that file is a zip file
2.  [X] Unzip file into temp dir
3.  [X] Check for and load [Content_Types].xml from root dir.  If it doesn't exist, it isn't an OOXML.
4.  [X] Verify OOXML by checking top of [Content_Types].xml: 
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
5.  [X] Load all "PartName" into array of "PartName" objects
6.  [ ] Search for "PartName" where "ContentType" in:
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" (Excel)
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" (Word)
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml" (PowerPoint)
7.  [ ] If we find that we are Excel, continue to load, otherwise exception
8.  [ ] Load "PartName" where "ContentType" = application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml
    (This should be "xl/workbook.xml")
    This is our "Excel" object
9.  [ ] In our Excel object, map "externalReferences" using "r:id"="id" to xl/_rels/workbook.xml.rels
10. [ ] Check "Target" of all references, split path and load externalLinks/_rels/<Target>.xml.rel


#>


    [System.IO.FileInfo]$file
    [tempDir]$tmpDir
    [xml]$RootXML
    [OOXMLParts]$OOXMLParts
    [boolean]$loaded = $false

    # Static Vars
    static [string]$RootXMLPath = '[Content_Types].xml'
    static [string]$OOXMLNS = 'http://schemas.openxmlformats.org/package/2006/content-types'


    <##
     # CONSTRUCTORS
     ##>
    OOXML(){
        # Nothing to do - File must be set and loaded later
    }
    OOXML([string]$filename){
        $this.setFilename($filename)
        $this.load()
    }


    <##
     # Set the filename of the file we are working on
     #
     #>
    [OOXML]setFilename([string]$filename){
        if($this.loaded){
            throw "Cannot load - already loaded!"
        }

        # Get-Item loads information about a file
        $fp = Get-Item $filename
        $this.file = $fp

        return $this
    }


    <##
     # Load the file we are working on into memory
     #
     # ...I would love to do this in actual memory.  In
     # reality we are doing it in a tmp folder.  :-(
     #>
    [OOXML]load(){
        if($this.loaded){
            throw "Cannot load - already loaded!"
        }

        # Create TempDir
        # TODO: See if we can work on OOXML on-the-fly rather than unzipping
        $this.tmpDir = [tempDir]::new()

        try{
            # Unzip XLSX into TempDir
            $tmpPath = $this.tmpDir.getPath()
            Write-Verbose "Extracting OOXML to $tmpPath"
            Expand-Archive -LiteralPath $this.file.FullName -DestinationPath $tmpPath

            # Load the root XML
            Write-Debug ("Loading OOXML: " + $this.getRootXmlPath())
            # Get-Content loads the content of the file
            [xml]$this.RootXML = Get-Content -LiteralPath $this.getRootXmlPath()
        }
        catch{
            $this.tmpDir.Dispose()
            throw "Not an OOXML file!"
        }

        if(!$this.checkOOXML()){
            $this.tmpDir.Dispose()
            throw "Not an OOXML file!"
        }

        # Load the OOXML Parts
        $this.OOXMLParts = [OOXMLParts]::new()
        foreach($part in $this.RootXML.Types.Override){
            $part = [OOXMLPart]::new($part.PartName, $part.ContentType)
            $this.OOXMLParts.Add($part)
        }

        $this.loaded = $true

        return $this
    }


    <#####################
    ##### SAVE LOGIC #####
    #####################>


    <##
     # Save (overwrite) the original OOXML
     #
     # @return System.IO.FileInfo Returns FileInfo object of saved file
    #>
    [System.IO.FileInfo]save(){
        $filename = $this.file.FullName
        return $this.doSave($filename, $true)
    }


    <##
     # Save OOXML file with a new filename
     #
     # @param string filename Filename to save as.  If no path given, current path will be used.
     # @param boolean overwrite Optional.  True if we want to overwrite an existing file.  Defaults to False
     # @return System.IO.FileInfo Returns the FileInfo object of the saved file
     ##>
    [System.IO.FileInfo]saveAs([string]$filename){
        # TODO: Need some logic to check for absolute or relative paths for filename
        return $this.saveAs($filename, $false)
    }
    [System.IO.FileInfo]saveAs([string]$filename, [boolean]$overwrite){
        # TODO: Need some logic to check for absolute or relative paths for filename
        return $this.doSave($filename, $overwrite)
    }


    <##
     # Actual save routine
     #
     # @param string filename *FULL PATH INCLUDING NAME* of file to save as
     # @param boolean overwrite True if we want to overwrite the existing file.  Defaults to False
     # @return System.IO.FileInfo Returns the FileInfo object of the saved file
     ##>
    hidden [System.IO.FileInfo]doSave([string]$filename, [boolean]$overwrite){
        # Check if we are loaded yet.  Error out if we aren't; We can't save what we haven't loaded!
        if(!$this.loaded){
            throw "Cannot save before OOXML has been loaded!"
        }
        if(Test-Path $filename){
            if($overwrite){
                try{
                    Remove-Item $filename
                }
                catch{
                    throw "Cannot remove existing file!"
                }
            }
            else{
                throw "File already exists!"
            }
        }


        # Save our object shit to the temp dir first
        $this.RootXML.Save($this.getRootXmlPath())


        # Zip up TempDir into $filename
        # Compress-Archive currently has a bug where there is no way of zipping hidden files.
        # This is a problem because some of the files are dotfiles
        #Compress-Archive -Force -Path (Join-Path $this.tmpDir.getPath() "*" ) -DestinationPath $filename

        # Workaround: https://stackoverflow.com/a/70608133
        try{
            [System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $this.tmpDir.getPath() "/" ), $filename)
        }
        catch{
            throw "Failed to create file!"
        }

        $fp = Get-Item $filename

        $this.file = $fp
        return $this.file
    }


    <#########################
    ##### END SAVE LOGIC #####
    #########################>

    <##################
    ##### GETTERS #####
    ##################>


    <##
     # Gets an XML object starting at $node
     #
     # @param string Specifies an XPath search query. The query language is case-sensitive.
     # @return xml Returns an XML object matching XPath
     #>
    [xml]getNode([string]$XPath){
        $xml = Select-Xml -Content $this.RootXML -XPath $XPath
        return $xml
    }


    <##
     # Get filename we are currently working on
     #>
    [string]getFilename(){
        return $this.file.name
    }


    <##
     #
     #>
    [string]getRootXmlPath(){
        return (Join-Path $this.tmpDir.getPath() $this::RootXMLPath)
    }

    <##
     # @return The root path of our temp dir where we are working on the OOXML
     #>
    [string]getRootPath(){
        return $this.tmpDir.getPath()
    }

    <######################
    ##### END GETTERS #####
    ######################>


	[boolean]checkOOXML(){
		if($this.RootXML.Types.xmlns -eq [OOXML]::OOXMLNS){
			return $true
		}
		else{
			return $false
		}
	}


    <##
	 # toString magic method
	 # This method is called when an object is called as a string
	 #
	 # @return string Returns a string in the format of "<Filename>"
	 #>
	[string]toString(){
		$str = ""
		$str += $this.filename
		return $str
	}


    <##
     # Implement Dispose to fullfill the IDisposable interface
    #>
    Dispose(){
        if(!$this.loaded){
            Write-Debug "Not loaded yet - nothing to dispose"
        }
        # Nuke the temp directory and everything in it.
        $this.tmpDir.Dispose()
        $this.loaded = $false
        #base.Dispose  # This is supposed to be called, but errors out
    }


}
