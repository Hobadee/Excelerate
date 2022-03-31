class XLSX : OOXML {

    [xml]$workbook

    # Static Vars
    static [string]$XLSXPartType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'


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
        if(([OOXML]$this).load().checkXLSX() -ne $true){
            $this.tmpDir.Dispose()
            throw "Not an XLSX file!"
        }
        else{
            return $this
        }
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
