<#
.SYNOPSIS
XMLDocument that stores a Stream reference so it can save itself later

.DESCRIPTION
We inherit the [System.Xml.XmlDocument] class, and implement [IDisposable]
#>
class XmlPart : System.Xml.XmlDocument, IDisposable{
    [System.IO.Stream]$stream
    [System.IO.Packaging.ZipPackagePart]$part


    <#
    .PARAMETER $part
    ZipPackagePart to be loaded as an XML file

    .PARAMETER $FileMode
    FileMode we should load this part as.

    .LINK
    https://learn.microsoft.com/en-us/dotnet/api/system.io.filemode

    .PARAMETER $FileAccess
    FileAccess we should load this part as

    .LINK
    https://learn.microsoft.com/en-us/dotnet/api/system.io.fileaccess

    #>
    XmlPart([System.IO.Packaging.ZipPackagePart]$part, [System.IO.FileMode]$FileMode, [System.IO.FileAccess]$FileAccess) : base(){
        $this.part = $part
        $this.stream = $this.part.GetStream([OOXML]::FileMode, [OOXML]::FileAccess)
        $this.Load()
    }


    <#
    I shouldn't enable this until I have better error handling in this class
    #>
    <#
    setStream($stream){
        $this.stream = $stream
    }
    #>


    <#
    .SYNOPSIS
    Returns the stream object
    #>
    [System.IO.Stream] getStream(){
        return $this.stream
    }


    <#
    .SYNOPSIS
    Returns the relationships this XML has
    #>
    [System.IO.Packaging.PackageRelationshipCollection]GetRelationships(){
        return $this.part.GetRelationships()
    }


    <#
    .SYNOPSIS
    Returns the package-level relationship with a given identifier.

    .PARAMETER id
    The Id of the relationship to return.

    .OUTPUTS
    The package-level relationship with the specified id.
    #>
    [System.IO.Packaging.PackageRelationship]GetRelationship([string]$id){
        return $this.part.GetRelationship($id)
    }


    <#
    .SYNOPSIS
    Overload Load() method to take existing stream into account
    #>
    [XmlPart]Load(){
        ([System.Xml.XmlDocument]$this).Load($this.stream)
        return $this
    }


    <#
    .SYNOPSIS
    Overload Save() method to take existing stream into account
    #>
    [XmlPart]Save(){
        $this.Save($this.stream)
        return $this
    }


    <#
    .SYNOPSIS
    Implement Dispose to fullfill the IDisposable interface

    .DESCRIPTION
    This closes the open stream
    #>
    Dispose(){
        $this.stream.Dispose()
    }

}
