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


    <##################
    ##### LOADERS #####
    ##################>
    <#
     # Stuff here is just part of [XLSX]load(), but would be too messy in a single method
     #>


    hidden[boolean]loadWorkbook(){
        $tmpPath = $this.tmpDir.getPath()
        $wrkbk = (Join-Path $tmpPath $this.OOXMLParts.getPartByType($this::XLSXPartType).getName())

        Write-Debug "Loading XLSX Workbook: $wrkbk"
        # Get-Content loads the content of the file
        [xml]$this.workbook = Get-Content -LiteralPath $wrkbk

        if($this.workbook.workbook.xmlns -ne $this::XLSXWorkbookXMLNS){
            $this.tmpDir.Dispose()
            throw "Invalid XLSX file!"
        }
        return $true
    }


    hidden[boolean]loadDefinedNames(){
        return $true
    }


    hidden[boolean]loadExternalLinks(){
        return $true
    }


    <########################
    ##### END OVERRIDES #####
    ########################>


	[boolean]checkXLSX(){
		if(([OOXML]$this).OOXMLParts.getPartByType($this::XLSXPartType) -ne $null){
			return $true
		}
		else{
			return $false
		}
	}

}
