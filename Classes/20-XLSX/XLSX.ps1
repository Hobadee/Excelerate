class XLSX : OOXML {

    hidden [XmlStream] $workbook
    hidden [XmlStream] $relations
    hidden [XLSXExternalLinks] $XLSXExternalLinks = [XLSXExternalLinks]::new()

    # Static Vars
    static [string] $XLSXWorkbookXMLNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    static [string] $XLSXWorkbook = "/xl/workbook.xml"
    static [string] $XLSXRelations = "/xl/_rels/workbook.xml.rels"


    <#################################
     #          Constructor          #
     #################################>


    XLSX([string]$Filename) : base($Filename){}


    <###############################
     #          Overrides          #
     ###############################>


    [XLSX]load(){
        ([OOXML]$this).load()

        if(!$this.isValid()){
            [Logging]::Debug("Couldn't verify XLSX data")
            throw [System.IO.FileLoadException]::new()
        }

        # Load master workbook
        $this.workbook = $this.getXmlFile([XLSX]::XLSXWorkbook)

        # Load the Relations XML
        $this.relations = $this.getXmlFile([XLSX]::XLSXRelations)

        # Load external links
        [Logging]::Debug("Loading External Links...")
        foreach($ref in $this.workbook.DocumentElement.externalReferences.ChildNodes){
            $link = [XLSXExternalLink]::new($ref, $this)
            $this.XLSXExternalLinks.Add($link)
        }

        return $this
    }


    <##############################
     #          Checkers          #
     ##############################>


    <#
    .SYNOPSIS
    Override method to check if this is a valid XLSX file or not
    #>
    [Boolean]isValid(){
        if(!$this.isLoaded()){
            throw [NullReferenceException]::new()
        }
        # This file contains the core properties for any Excel XML document.
        return $this.OOXML.PartExists([XLSX]::XLSXWorkbook)
    }


    <#####################################
     #          File Operations          #
     #####################################>


    <#
    .SYNOPSIS
    Overload the parent doSave method so we can save our XLSX XML first before OOXML zips it up
    #>
    hidden [System.IO.FileInfo]doSave([string]$filename, [boolean]$overwrite){
        # Save logic for XLSX stuff
        # Make sure we save our XML!
        $this.workbook.Save($this.getWorkbookXmlPath())

        # Once our XLSX saving is done, call parent method
        return ([OOXML]$this).doSave($filename, $overwrite)
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
        $this.workbook.save()
        ([OOXML]$this).save()
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

        $this.workbook.Dispose()
        ([OOXML]$this).Dispose()
    }


    <#############################
     #          Getters          #
     #############################>


    getWorkbookXmlPath(){
        # Dead function - DO NOT IMPLEMENT!
        # Still references to it in code, so here for compiling reasons
        [Logging]::Debug("getWorkbookXmlPath() called.  Let me die in peace!")
        throw [NotImplementedException]::new()
    }


    <###################################
     #          Misc Bullshit          #
     ###################################>


    <#
    .SYNOPSIS
    Remove all broken refs

    .DESCRIPTION
    Remove all broken refs.  In original design, refs were searched first, then removed.  This
    was a longer process, but would allow for manual intervention in the future.  We should consider
    going back to this design in the future, but for now a blanked removal will suffice.
    #>
    removeBrokenRefs(){
        $removed = 0;
        foreach($name in $this.workbook.DocumentElement.definedNames.ChildNodes){
            if($name.'#text' -like '*#REF!*'){
                #$name.RemoveAll()
                $removed++
            }
        }
        [Logging]::Debug("$removed broken references removed")
    }


    <##
     # This should *NOT* be in the final version!
     # This is a temporary hack to get us to MVP quickly
     # Since parsing ExternalReferences will be a major project,
     # and often we can just nuke them all, enable nuking of all ExternalReferences.
     #
     # TODO: Ensure this can be implemented cleanly via child methods.  This may
     #   involve a few different method calls
     #>
    [XLSX]removeAllExternalReferences(){
        Write-Verbose "Removing all External References!"
        $this.workbook.DocumentElement.externalReferences.RemoveAll()
        Remove-Item -Recurse (Join-Path $this.getRootPath() "xl" "externalLinks")
        return $this
    }

}
