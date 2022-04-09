class XLSX : OOXML {

    [xml]$workbook
    [XLSXDefinedNames]$XLSXDefinedNames = [XLSXDefinedNames]::new()
    [XLSXExternalLinks]$XLSXExternalLinks = [XLSXExternalLinks]::new()

    # Static Vars
    static [string]$XLSXPartType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
    static [string]$XLSXWorkbookXMLNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


    <#######################
    ##### CONSTRUCTORS #####
    #######################>


    XLSX(){
        # Nothing to do - File must be set and loaded later
    }
    XLSX([string]$filename){
        $this.setFilename($filename)
        $this.load()
    }


    <###########################
    ##### END CONSTRUCTORS #####
    ###########################>


    <####################
    ##### OVERRIDES #####
    ####################>


    [XLSX]load(){
        ([OOXML]$this).load()
        if($this.checkXLSX() -ne $true){
            $this.tmpDir.Dispose()
            throw "Not an XLSX file!"
        }

        $this.loadWorkbook()
        $this.loadDefinedNames()
        $this.loadExternalLinks()

        return $this
    }


    <########################
    ##### END OVERRIDES #####
    ########################>


    <##################
    ##### LOADERS #####
    ##################>
    <#
     # Stuff here is just part of [XLSX]load(), but would be too messy in a single method
     #>


    hidden[boolean]loadWorkbook(){
        Write-Debug ("Loading XLSX Workbook: " + $this.getWorkbookXmlPath())
        # Get-Content loads the content of the file
        [xml]$this.workbook = Get-Content -LiteralPath $this.getWorkbookXmlPath()

        if($this.workbook.workbook.xmlns -ne $this::XLSXWorkbookXMLNS){
            $this.tmpDir.Dispose()
            throw "Invalid XLSX file!"
        }
        return $true
    }


    hidden[boolean]loadDefinedNames(){
        Write-Debug "Loading Defined Names..."
        $i = 0
        foreach($item in $this.workbook.workbook.definedNames.definedName){
            $i++
            $name = [XLSXDefinedName]::new()
            $name.setName($item.name)
            $name.setValue($item.'#text')
            if($item.localSheetId -ne $null){
                $name.setLocalSheetId($item.localSheetId)
            }
            if($item.hidden){
                $name.setHidden($true)
            }
            $this.XLSXDefinedNames.Add($name)
        }
        Write-Debug "$i Defined Names loaded."
        return $true
    }


    hidden[boolean]loadExternalLinks(){
        return $true
    }


    <######################
    ##### END LOADERS #####
    ######################>

    <###############
    ##### SAVE #####
    ###############>


    <##
     # Overload the parent doSave method so we can save our XLSX XML first before OOXML zips it up
     #>
    hidden [System.IO.FileInfo]doSave([string]$filename, [boolean]$overwrite){
        # Save logic for XLSX stuff
        # Make sure we save our XML!
        $this.workbook.Save($this.getWorkbookXmlPath())

        # Once our XLSX saving is done, call parent method
        return ([OOXML]$this).doSave($filename, $overwrite)
    }


    <###################
    ##### END SAVE #####
    ###################>


    <##################
    ##### GETTERS #####
    ##################>


    <##
     # Get the path of the main XLSX Workbook XML
     #>
    [string]getWorkbookXmlPath(){
        return (Join-Path $this.tmpDir.getPath() $this.OOXMLParts.getPartByType($this::XLSXPartType).getName())
    }


    <######################
    ##### END GETTERS #####
    ######################>
    


	[boolean]checkXLSX(){
		if(([OOXML]$this).OOXMLParts.getPartByType($this::XLSXPartType) -ne $null){
			return $true
		}
		else{
			return $false
		}
	}

}
