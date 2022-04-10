<#

For External Links, read `rId` from `xl/workbook.xml` and cross reference to
`<Relationship Id="rId##">` in `xl/_rels/workbook.xml` to get the `Target`.

We load `Target` and get a bunch of (useless?) sheet data.  We load
`Target/_rels/Target.rels` and get a new `Target` (URI) which we can test for
connectivity.  If no connectivity, delete all references.

As part of cleanup, ensure that neither is orphaned and removed if orphaned.

#>


class XLSXExternalLink{


    [string]$rId
    [System.IO.FileInfo]$targetFile
    [string]$targetURI
    [boolean]$UriTested = $false
    [boolean]$UriWorks = $null


    XLSXExternalLink(){
        # Nothing to do
    }


    <##
     # Tests if the links URI works or not.
     # Will only test once - subsequent runs will return cached value
     #>
    [boolean]testURI(){
        if($this.UriTested -eq $true){
            return $this.UriWorks
        }


        #TODO: Implement me!
        <#
        # Proper testing, especially cross-platform, will be hard
        # Multiple protocols must be taken into account such as
        # SMB, HTTP(S), and local filesystem
        #
        # There doesn't appear to be an "easy" button for this.
        # [System.Uri] class could at least help parse things out.  Check C# classes for more possibilities
        # Might need to make our own class that does a switch() for various types of URIs
        #
        # For files, Test-Path works
        # For web URLs, Invoke-Webrequest might be able to help us.
        # 
        #>


        $this.UriTested = $true

        return $this.UriWorks
    }


    <##
     # Invalidate the cached or manual test result and force a retest of the URI
     #
     # NOTE: This will override a manually set URI status!
     #
     # @return boolean $true if URI works, $false if not
     #>
    [boolean]forceTestURI(){
        $this.UriTested = $false
        return $this.testURI()
    }


    <##
     # Forces the URI check to a manual value
     # Useful for manually overriding URIs before later removal of non-working URIs
     #
     # $var boolean Status of the URI in question.  $true = working, $false = nonworking
     # @return XLSXExternalLink Returns this object
     #>
    [XLSXExternalLink]forceURIStatus([boolean]$status){
        $this.UriTested = $true
        $this.UriWorks = $status
        return $this
    }


    <##################
    ##### GETTERS #####
    ##################>
    [string]getRID(){
        return $this.rId
    }
    [System.IO.FileInfo]getTargetFile(){
        return $this.targetFile
    }
    [System.IO.FileInfo]getTargetRelFile(){
        # This won't actually work.  Still need to fully dev it
        return (Join-Path $this.targetFile "_rels" ($this.targetFile + ".rels"))
    }
    [string]getTargetURI(){
        return $this.targetURI
    }


    <##################
    ##### SETTERS #####
    ##################>
    [XLSXExternalLink]setRID([string]$RID){
        $this.rId = $RID
        return $this
    }
    [XLSXExternalLink]setTargetFile([string]$targetFile){
        $this.targetFile = $targetFile
        return $this
    }
    [XLSXExternalLink]setTargetURI([string]$targetURI){
        $this.targetURI = $targetURI
        return $this
    }


}
