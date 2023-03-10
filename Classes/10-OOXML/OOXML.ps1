

<#
.SYNOPSIS
Class to load and handle OOXML files

.DESCRIPTION
Class to load and handle OOXML files.  Largely a wrapper for the System.IO.Packaging.ZipPackage class, designed to be extended
with child classes, based on the exact type of file.


#>
class OOXML : IDisposable {

<#
# Flow
1.  [X] Attempt to load object as System.IO.Packaging.Package object
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

    # OOXML as ZIP file object in RAM
    hidden [System.IO.Packaging.ZipPackage] $OOXML

    hidden [String]$Filename
    hidden [System.Xml.XmlDocument]$RootXML

    # Static Vars
    hidden static [string] $OOXMLCore = "/docProps/core.xml"
    hidden static $FileAccess = [System.IO.FileAccess]::ReadWrite
    hidden static $FileMode = [System.IO.FileMode]::Open


    <#
    .SYNOPSIS
    Class constructor

    .PARAMETER Filename
    Filename of the OOXML file to load

    .EXAMPLE
    OOXML::new()

    .EXAMPLE
    OOXML::new([String] $Filename)

    #>
    OOXML([String]$Filename){
        $this.Filename = $filename
        $this.load()
    }
    

    <#############################
     #          Getters          #
     #############################>


    <#
    .SYNOPSIS
    Returns an XML object of the XML file requested

    .OUTPUTS
    XmlStream.  XmlStream object.  (This just extends System.Xml.XmlDocument to easily load/save)
    #>
    [XmlStream] getXmlFile([String] $Filename){
        #$stream = $this.getFileStream($Filename)


        <#
        [Logging]::Debug("Temp disable stuff - clean up later")
        $part = $this.OOXML.GetPart($Filename)
        [Logging]::Debug("Filename:"+$Filename+"\nPart:"+$part)
        [Logging]::Debug($part.GetRelationships())
        $stream = $part.GetStream([System.IO.FileMode]::Open)
        $xml = [XmlStream]::new($stream)
        #>


        $stream = $this.OOXML.GetPart($Filename).GetStream([OOXML]::FileMode, [OOXML]::FileAccess)
        $xml = [XmlStream]::new($stream)
        return $xml
    }


    <#
    .SYNOPSIS
    Returns a filestream object of the file requested
    #>
    [System.IO.Stream] getFileStream([String] $Filename){
        try{
            $stream = $this.OOXML.GetPart($Filename).GetStream([OOXML]::FileMode, [OOXML]::FileAccess)
        }
        catch{
            throw [System.IO.FileNotFoundException]::new()
        }
        return $stream
    }


    <#
    .SYNOPSIS
    Get filename we are currently working on
    #>
    [String] getFilename(){
        return $this.filename
    }


    <##############################
     #          Checkers          #
     ##############################>


    <#
    .SYNOPSIS
    Checks if the OOXML object is loaded in memory

    .OUTPUTS
    Boolean
    True if OOXML is loaded in memory
    False otherwise
    #>
    [Boolean]isLoaded(){
        if(!$this.OOXML){
            return $False
        }
        return ($this.OOXML.FileOpenAccess -eq [OOXML]::FileAccess)
    }


    <#
    .SYNOPSIS
    Checks if the file is valid OOXML

    .DESCRIPTION
    Checks if the loaded file is valid OOXML.

    Child methods should override this and do their own checks for their respective child types.
    #>
    [Boolean]isValid(){
        if(!$this.isLoaded()){
            throw [NullReferenceException]::new()
        }
        # https://en.wikipedia.org/wiki/Office_Open_XML_file_formats
        # This file contains the core properties for any Office Open XML document.
        return $this.OOXML.PartExists([OOXML]::OOXMLCore)
    }


    <#####################################
     #          File Operations          #
     #####################################>


    <#
    .SYNOPSIS
    Load the file we are working on into memory
    #>
    [OOXML]load(){

        if($this.isLoaded()){
            throw [System.InvalidOperationException]::new()
        }

        # Attempt to load file as ZIP object into memory.  If we can't load as a ZIP, we aren't OOXML
        try{
            $this.OOXML = [System.IO.Packaging.Package]::Open($this.Filename, [OOXML]::FileMode, [OOXML]::FileAccess)
        }
        catch{
            Write-Debug("Couldn't load ZIP file object")
            [Logging]::trace($_)
            throw [System.IO.FileLoadException]::new()
        }

        # Do a real check if we are actual OOXML and puke if not
        if(!$this.isValid()){
            Write-Debug("Couldn't verify OOXML data")
            throw [System.IO.FileLoadException]::new()
        }

        return $this
    }


    <#
    .SYNOPSIS
    Save (overwrite) the original OOXML

    .DESCRIPTION
    This saves (overwriting) the file, but does NOT close it.
    #>
    save(){
        if(!$this.isLoaded()){
            throw [NullReferenceException]::new()
        }
        $this.OOXML.flush()
    }


    <#
    .SYNOPSIS
    Implement Dispose to fullfill the IDisposable interface

    .DESCRIPTION
    This saves (overwriting) and closes the file
    #>
    Dispose(){
        if(!$this.isLoaded()){
            throw [NullReferenceException]::new()
        }
        # NOTE: Dispose on the OOXML object saves it!
        $this.OOXML.Dispose()
        #base.Dispose  # This is supposed to be called, but errors out
        $this.OOXML = $null
    }


    <###################################
     #          Magic Methods          #
     ###################################>


    <#
    .SYNOPSIS
	toString() magic method

    .DESCRIPTION
	This method is called when an object is called as a string

	.OUTPUTS
	string. Returns a string in the format of "<Filename>"
	#>
	[String]toString(){
		$str = ""
		$str += $this.Filename
		return $str
	}


}
