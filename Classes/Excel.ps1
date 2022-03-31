class Excel : IDisposable {
    
    [System.IO.FileInfo]$file
    [tempDir]$tmpDir
    [xml]$workbook
    [boolean]$loaded = $false

    # Static Vars
    static [string]$workbookFile = 'xl/workbook.xml'


    <##
     # CONSTRUCTORS
     ##>
    Excel(){
        # Nothing to do - File must be set and loaded later
    }
    Excel([string]$filename){
        $this.setFilename($filename)
        $this.load()
    }


    [boolean]setFilename([string]$filename){
        if($this.loaded){
            throw "Cannot load - already loaded!"
        }

        # Get-Item loads information about a file
        $fp = Get-Item $filename
        if($fp.Extension -ne ".xlsx"){
            throw "Unknown file type!"
        }

        $this.file = $fp

        return $true
    }


    [Excel]load(){
        if($this.loaded){
            throw "Cannot load - already loaded!"
        }

        # Create TempDir
        # TODO: See if we can work on XLSX on-the-fly rather than unzipping
        $this.tmpDir = [tempDir]::new()

        # Unzip XLSX into TempDir
        $tmpPath = $this.tmpDir.getPath()
        Write-Verbose "Extracting XLSX to $tmpPath"
        Expand-Archive -LiteralPath $this.file.FullName -DestinationPath $tmpPath

        $xlsxWorkbook = (Join-Path $tmpPath $this::workbookFile)
        Write-Debug "Loading Workbook: $xlsxWorkbook"
        # Get-Content loads the content of the file
        [xml]$this.workbook = Get-Content -Path $xlsxWorkbook

        $this.loaded = $true

        return $this
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
     # Gets an XML object starting at $node
     #
     #
     #
     #>
    [xml]getNode([string]$XPath){
        $xml = Select-Xml -Content $this.workbook -XPath $XPath
        return $xml
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
        $fp = Get-Item $path
        return $fp
     }


    <##
     # Save (overwrite) the original XLSX
     #
     # @return System.IO.FileInfo Returns FileInfo object of saved file
    #>
    [System.IO.FileInfo]save(){
        $filename = $this.file.FullName
        return $this.doSave($filename, $true)
    }


    <##
     # Save XLSX file with a new filename
     #
     # @var string filename Filename to save as.  If no path given, current path will be used.
     # @var boolean overwrite Optional.  True if we want to overwrite an existing file.  Defaults to False
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
     # @var string filename *FULL PATH INCLUDING NAME* of file to save as
     # @var boolean overwrite True if we want to overwrite the existing file.  Defaults to False
     # @return System.IO.FileInfo Returns the FileInfo object of the saved file
     ##>
    hidden [System.IO.FileInfo]doSave([string]$filename, [boolean]$overwrite){
        # Check if we are loaded yet.  Error out if we aren't; We can't save what we haven't loaded!
        if(!$this.loaded){
            throw "Cannot save before XLSX has been loaded!"
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

        # Zip up TempDir into $filename
        # Compress-Archive currently has a bug where there is no way of zipping hidden files.
        # This is a problem because some of the files are dotfiles
        #Compress-Archive -Force -Path (Join-Path $this.tmpDir.getPath() "*" ) -DestinationPath $filename

        # Workaround: https://stackoverflow.com/a/70608133
        [System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $this.tmpDir.getPath() "/" ), $filename)

        $fp = Get-Item $filename

        $this.file = $fp
        return $this.file
    }


    <##
     # Implement Dispose to fullfill the IDisposable interface
    #>
    Dispose(){
        if(!$this.loaded){
            Write.Debug "Not loaded yet - nothing to dispose"
        }
        # Nuke the temp directory and everything in it.
        $this.tmpDir.Dispose()
        $this.loaded = $false
        #base.Dispose  # This is supposed to be called, but errors out
    }
}
