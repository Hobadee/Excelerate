class OldExcel : IDisposable {
    [string]$filename
    [tempDir]$tmpDir
    [boolean]$loaded = $false
    [xml]$workbook
    [string]$workbookFile = "xl/workbook.xml"

    # I might want to be using System.IO.FileInfo to track things, or at least extend that class


    OldExcel(){
        # Empty Constructor
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
     # Function to load the Excel file for processing
     #
     # @return Excel Returns self on success (Throws error on failure)
    #>
    [Excel]load(){

        # Check if we are loaded already.  Error out if we are; We can't load twice!
        if($this.loaded){
            throw "Cannot load:  XLSX alread loaded!"
        }
        else{
            Write-Debug "Excel not loaded: Loading"
            $this.loaded = $true
        }

        if($this.filename -eq $null){
            throw "Cannot load:  No file selected!"
        }

        # TODO: Create TempDir
        $this.tmpDir = [tempDir]::new()
        
        # TODO: Unzip XLSX into TempDir
        $xlsxPath = $this.tmpDir.getPath()
        Write-Verbose "Extracting XLSX to $xlsxPath"
        Expand-Archive -LiteralPath $this.filename -DestinationPath $xlsxPath

        $xlsxWorkbook = (Join-Path $xlsxPath $this.workbookFile)
        Write-Debug "Loading Workbook: $xlsxWorkbook"
        [xml]$this.workbook = Get-Content -Path $xlsxWorkbook

        return $this
    }


    <##
     # Function to return a file object for a file inside the XLSX file
     #
     # @var string Path to file (relative to root of XLSX)
     # @return System.IO.FileInfo File object of enclosed file
     ##>
     [System.IO.FileInfo]getFile([string]$path){
        $path = (Join-Path $this.getTempDir $path)
        $file = Get-ChildItem $path
        return $file
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
            throw "Cannot change load after XLSX has been loaded!"
        }

        # TODO: Zip up TempDir into Filename


        # TODO: Remove TempDir
        $this.tmpDir.Dispose

        # We are saved & cleaned up; Unset our loaded variable
        $this.loaded = $false

        return $true
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
