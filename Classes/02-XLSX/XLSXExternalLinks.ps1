class XLSXExternalLinks : System.Collections.Generic.List[PSObject]{


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
     # Removes *NON WORKING* links from this object
     # 
     # NOTE: This is desctructive and will remove data!
     #
     # NOTE: This method may take some time as it will test all URIs if
     # they haven't been tested yet.
     #
     # @return int Returns number of links removed
     #>
    [int]removeBrokenLinks(){
        [int]$removed = 0
        
        foreach($XLSXExternalLink in $this){
            if($XLSXExternalLink.testURI() -eq $false){
                if($this.Remove($XLSXExternalLink)){
                    $removed++
                }
                else{
                    $str = "Failed to remove item: " + $XLSXExternalLink.getRID()
                    throw $str
                }
            }
        }
        return $this
    }


}
