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
     # Find invalid links in this object and return a new object containing those broken link objects
     #
     # NOTE: This may take quite some time to run if `testURI()` hasn't been run on the children yet.
     #
     # @return XLSXExternalLinks Returns a NEW XLSXExternalLinks object filled with broken link objects
     #>
    [XLSXExternalLinks]findBrokenLinks(){

        # Set up vars we need
        $links = [XLSXExternalLinks]::new()
        [int]$i = 0
        $total = $this.count
        
        foreach($item in $this){
            if($item.testURI() -eq $false){
                $links.add($item)
            }

            # Since this can run slow, give a progress bar
            $i++
            $pct = [math]::Round((($i / $total)*100),2)
            Write-Progress -Activity "Finding broken links" -Status "$pct% Complete" -PercentComplete $pct
        }

        Write-Progress -Activity "Finding broken links" -Completed
        return $links
    }


    <##
     # Removes links from this XLSXExternalLinks object by passing them as a new XLSXExternalLinks object
     #
     # NOTE: This is desctructive and will remove data!
     #
     # NOTE: This may take some time to run on large datasets
     #
     # @param XLSXExternalLinks $names Takes a XLSXExternalLinks list of XLSXExternalLink objects to remove from this object
     # @return int Number of items removed from this object
     #>
    [int]removeLinksByObjects([XLSXExternalLinks]$links){

        # Set up vars we need
        [int]$removed = 0
        [int]$i = 0
        $total = $links.count

        foreach($item in $links){
            if($this.Remove($item)){
                $removed++
            }
            else{
                $str = "Failed to remove item: " + $item.getName()
                throw $str
            }

            # Since this can run slow, give a progress bar
            $i++
            $pct = [math]::Round((($i / $total)*100),2)
            Write-Progress -Activity "Removing links" -Status "$pct% Complete: $i of $total removed" -PercentComplete $pct
        }

        Write-Progress -Activity "Removing links" -Completed
        return $removed
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
        $links = $this.findBrokenLinks()
        $removed = $this.removeLinksByObjects($links)
        return $removed
    }


}
