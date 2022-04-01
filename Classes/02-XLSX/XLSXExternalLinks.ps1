class XLSXExternalLinks : System.Collections.Generic.List[PSObject]{


    [boolean]$nested = $false


    <#
     # We should probably have a method to check for duplicate rIds...?
     #>


    <#
     # Returns the *FIRST* link matching $rId
     #>
    [XLSXExternalLink]getLinkByRid([string]$rId){
        foreach($XLSXExternalLink in $this){
            if($XLSXExternalLink.getRID() -eq $rId){
                return $XLSXExternalLink
            }
        }
        return $null
    }


    <##
     # Returns an [XLSXExternalLinks] object of *NON WORKING* links
     # This is useful for later removal.
     #
     # NOTE: This method may take some time as it will test all URIs if
     # they haven't been tested yet.
     #
     # @return XLSXExternalLinks Object with list of link objects that are broken
     #>
    [XLSXExternalLinks]getBrokenLinks(){
        if($this.nested -eq $true){
            throw "Nesting [XLSXExternalLinks]::getBrokenLinks() is retarded"
        }

        $links = [XLSXExternalLinks]::new()
        $links.nested = $true
        
        foreach($XLSXExternalLink in $this){
            if($XLSXExternalLink.testURI() -eq $false){
                $links.add($XLSXExternalLink)
            }
        }
        return $links
    }


}
