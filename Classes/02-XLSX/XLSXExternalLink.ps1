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
        # Proper testing, especially cross-platform, will be hard
        # Multiple protocols must be taken into account such as
        # SMB, HTTP(S), and local filesystem


        $this.UriTested = $true

        return $this.UriWorks
    }


    <##
     # Invalidate the cached test result and force a retest of the URI
     #>
    [boolean]forceTestURI(){
        $this.UriTested = $false
        return $this.testURI()
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
