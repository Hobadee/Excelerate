<#
.SYNOPSIS
XMLDocument that stores a Stream reference so it can save itself later

.DESCRIPTION
We inherit the [System.Xml.XmlDocument] class, and implement [IDisposable]
#>
class XmlPart : System.Xml.XmlDocument, IDisposable{
    [System.IO.Stream]$stream
    [System.IO.Packaging.ZipPackagePart]$part


    XmlPart($stream, $part) : base(){
        $this.stream = $stream
        $this.part = $part
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
    Overload Load() method to take existing stream into account
    #>
    Load(){
        $this.Load($this.stream)
    }


    <#
    .SYNOPSIS
    Overload Save() method to take existing stream into account
    #>
    Save(){
        $this.Save($this.stream)
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
