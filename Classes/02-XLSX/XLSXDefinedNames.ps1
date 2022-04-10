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
        $removed = $this.removeNamesByObject($names)
        return $removed
    }


    <##
     # Find invalid references in this object, and return a new object containing those references
     # An invalid reference is any reference that contains "#REF!" ( -like "*#REF!*" )
     #
     # NOTE: This may take some time to run on large datasets
     #
     # @return XLSXDefinedNames Returns an XLSXDefinedNames object with a bunch of children XLSXDefinedName objects that have broken refs
     #>
    [XLSXDefinedNames]findBrokenRefs(){
        $names = [XLSXDefinedNames]::new()
        
        foreach($item in $this){
            if($item.getValue() -like '*#REF!*'){
                $names.add($item)
            }
        }

        return $names
    }


    <##
     # Removes names from this XLSXDefinedNames object by passing them as a new XLSXDefinedNames object
     #
     # NOTE: This is desctructive and will remove data!
     #
     # @param XLSXDefinedNames $names Takes a XLSXDefinedNames list of XLSXDefinedName objects to remove from this object
     # @return int Number of items removed from this object
     #>
    [int]removeNamesByObject([XLSXDefinedNames]$names){
        [int]$removed = 0
        foreach($item in $names){
            if($this.Remove($item)){
                $removed++
            }
            else{
                $str = "Failed to remove item: " + $item.getName()
                throw $str
            }
        }
        return $removed
    }

}
