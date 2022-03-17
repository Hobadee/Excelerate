class tempDir : IDisposable {    
    [string]$tempPath
    [string]$tempName


    <##
     # Constructor
     # Create a temporary directory to use elsewhere
    #>
    tempDir(){
        # Stolen from https://stackoverflow.com/a/34559554

        # Get the Temp Path from the OS
        $this.tempPath = [System.IO.Path]::GetTempPath()
        # Cast a GUID to a string and store it as our temp name
        [string]$this.tempName = [System.Guid]::NewGuid()

        if( -Not (Test-Path $this.tempPath) ){
            throw "Temp parent directory does not exist!"
        }
        if( Test-Path $this.getPath ){
            # Since our temp-dir is based off a GUID, this shouldn't ever happen
            # But who knows, you might be the lucky one!
            throw "Temp directory already exists!"
        }

        # Create the directory based off the information we have created
        New-Item -ItemType Directory -Path $this.getPath()

        Write-Verbose ('Temp dir created: ' + $this.getPath())
    }


    <##
     # Get the full path of the temp folder
     # @return string Full path of temp folder
    #>
    [string]getPath(){
        return [string](Join-Path $this.tempPath $this.tempName)
    }


    <##
     # Get the name of the temp folder
     # Useful if you already know the temp folder things go in
     # @return string Name of the temp folder
     #>
    [string]getName(){
        return $this.tempName
    }


    <##
     # Get the path of the temp dir's parent folder
     # @return string Name of the temp dir's parent folder
     #>
    [string]getParent(){
        return $this.tempPath
    }


    <##
     # Implement Dispose to fullfill the IDisposable interface
    #>
    Dispose(){
        # Nuke the temp directory and everything in it.

        # If Dispose() is called multiple times for some reason, Remove-Item will error out
        # Check if temp path still exists before removing.
        if( Test-Path $this.getPath() ){
            Remove-Item -Recurse -Force -LiteralPath $this.getPath()
            Write-Verbose ('Temp dir removed: ' + $this.getPath())
        }
        else{
            Write-Verbose ('Temp dir does not exist.  Cannot remove.')
        }
        #base.Dispose  # This is supposed to be called, but errors out
    }
}
