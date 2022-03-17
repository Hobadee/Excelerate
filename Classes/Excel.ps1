class Excel : IDisposable {
    
    [System.IO.FileInfo]$file
    [tempDir]$tmpDir
    [xml]$workbook
    [boolean]$loaded = $false
    [string]$workbookFile = "xl/workbook.xml"


    <##
     # CONSTRUCTOR
     ##>
    Excel([string]$filename){

        # Get-ChildItem loads information about a file
        # Get-Content loads the content of the file

        $fp = Get-ChildItem $filename
        if($fp.Extension -ne ".xlsx"){
            throw "Unknown file type!"
        }
        $this.file = $fp

        # Create TempDir
        # TODO: See if we can work on XLSX on-the-fly rather than unzipping
        $this.tmpDir = [tempDir]::new()
        
        # Unzip XLSX into TempDir
        $tmpPath = $this.tmpDir.getPath()
        Write-Verbose "Extracting XLSX to $tmpPath"
        Expand-Archive -LiteralPath $this.file.FullName -DestinationPath $tmpPath

        $xlsxWorkbook = (Join-Path $tmpPath $this.workbookFile)
        Write-Debug "Loading Workbook: $xlsxWorkbook"
        [xml]$this.workbook = Get-Content -Path $xlsxWorkbook

        $this.loaded = $true
        
    }


    <##
     # Returns the temporary directory the Excel file has been extracted to
     #
     # @return string The temporary directory the Excel file has been extracted to
    #>
    [string]getTempDir(){
        return $this.tmpDir.getPath()
    }


    [xml]getWorkbook(){
        return $this.workbook
    }


    [string]getFilename(){
        return $this.file.name
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
     # Checks if the Excel file is loaded or not
     # 
     # @return boolean TRUE if loaded, FALSE if unloaded
    #>
    [boolean]isLoaded(){
        return $this.loaded
    }


    <##
     # Function to set the Filename
     # If used before loading, will change what is loaded
     # If used after loading, will change the filename when saved
     #
     # @return Excel Returns self
    #>
    [Excel]setFilename([string]$filename){
        $this.filename = $filename
        return $this
    }


    <##
     # Check if the file is a valid Excel file
     # TODO: Implement me!
     #
     # @return boolean Returns true if the file is a valid Excel file, false otherwise
     ##>
    [boolean]checkValidity(){
        return $true
    }


    <##
     # Function to return a file object for a file inside the XLSX file
     #
     # @var string Path to file (relative to root of XLSX)
     # @return System.IO.FileInfo File object of enclosed file
     ##>
     [System.IO.FileInfo]getFile([string]$path){
        $path = (Join-Path $this.getTempDir $path)
        $fp = Get-ChildItem $path
        return $fp
     }


    <##
     #
     #
     #
     # @return boolean Returns true on success, false otherwise
    #>
    [boolean]save(){

        # Check if we are loaded yet.  Error out if we aren't; We can't save what we haven't loaded!
        if(!$this.loaded){
            throw "Cannot save before XLSX has been loaded!"
        }

        # TODO: Zip up TempDir into Filename


        # TODO: Remove TempDir
        $this.tmpDir.Dispose()

        # We are saved & cleaned up; Unset our loaded variable
        $this.loaded = $false

        return $true
    }


    <##
     # Save XLSX file with a new filename
     ##>
    [System.IO.FileInfo]saveAs($filename){
        # need to do some stuff here
        # $this.file = $filename
        # $this.save()
        throw "Not implemented yet!"
        return $this.file
    }


    <##
     # Implement Dispose to fullfill the IDisposable interface
    #>
    Dispose(){
        # Nuke the temp directory and everything in it.
        $this.tmpDir.Dispose()
        $this.loaded = $false
        #base.Dispose  # This is supposed to be called, but errors out
    }
}
