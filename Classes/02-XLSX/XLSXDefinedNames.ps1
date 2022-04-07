class XLSXDefinedNames : System.Collections.Generic.List[PSObject]{


    <##
     # Removes invalid references from this object
     #
     # NOTE: This will only remove references with the *EXACT* value of '#REF!'
     # 
     # NOTE: This is desctructive and will remove data!
     #
     # @return int Returns the number of references removed
     #>
    [int]removeBrokenRefs(){
        [int]$removed = 0
        foreach($XLSXDefinedName in $this){
            if($XLSXDefinedName.getValue() -eq '#REF!'){
                if($this.Remove($XLSXDefinedName)){
                    $removed++
                }
                else{
                    $str = "Failed to remove item: " + $XLSXDefinedName.getName()
                    throw $str
                }
            }
        }
        return $removed
    }

}
