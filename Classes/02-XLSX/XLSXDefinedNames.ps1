class XLSXDefinedNames : System.Collections.Generic.List[PSObject]{


    <##
     # Finds and Removes invalid references from this object
     # An invalid reference is any reference that contains "#REF!"
     #
     # NOTE: This may take some time to run on large datasets
     # 
     # NOTE: This is desctructive and will remove data!
     #
     # @return int Returns the number of references removed
     #>
    [int]removeBrokenRefs(){
        $names = $this.findBrokenRefs()
        $removed = $this.removeNamesByObjects($names)
        return $removed
    }


    <##
     # Find invalid references in this object, and return a new object containing those references
     # An invalid reference is any reference that contains "#REF!" ( -like "*#REF!*" )
     #
     # Logic to check for broken refs should probably be moved into XLSXDefinedName class later
     #
     # NOTE: This may take some time to run on large datasets
     #
     # @return XLSXDefinedNames Returns an XLSXDefinedNames object with a bunch of children XLSXDefinedName objects that have broken refs
     #>
    [XLSXDefinedNames]findBrokenRefs(){

        # Set up vars we need
        $names = [XLSXDefinedNames]::new()
        [int]$i = 0
        $total = $this.count
        
        foreach($item in $this){
            if($item.getValue() -like '*#REF!*'){
                $names.add($item)
            }

            # Since this can run slow, give a progress bar
            $i++
            $pct = [math]::Round((($i / $total)*100),2)
            Write-Progress -Activity "Finding broken references" -Status "$pct% Complete" -PercentComplete $pct
        }

        Write-Progress -Activity "Finding broken references" -Completed
        return $names
    }


    <##
     # Removes names from this XLSXDefinedNames object by passing them as a new XLSXDefinedNames object
     #
     # NOTE: This is desctructive and will remove data!
     #
     # NOTE: This may take some time to run on large datasets
     #
     # @param XLSXDefinedNames $names Takes a XLSXDefinedNames list of XLSXDefinedName objects to remove from this object
     # @return int Number of items removed from this object
     #>
    [int]removeNamesByObjects([XLSXDefinedNames]$names){

        # Set up vars we need
        [int]$removed = 0
        [int]$i = 0
        $total = $names.count

        foreach($item in $names){
            if($this.Remove($item)){
                Write-Verbose ("Removed defined name: " + $item.getName())
                $removed++
            }
            else{
                $str = "Failed to remove item: " + $item.getName()
                throw $str
            }

            # Since this can run slow, give a progress bar
            $i++
            $pct = [math]::Round((($i / $total)*100),2)
            Write-Progress -Activity "Removing defined names" -Status "$pct% Complete: $i of $total removed" -PercentComplete $pct
        }

        Write-Progress -Activity "Removing defined names" -Completed
        return $removed
    }

}
