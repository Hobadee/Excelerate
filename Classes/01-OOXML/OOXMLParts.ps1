class OOXMLParts : System.Collections.Generic.List[PSObject]{

    <#
     # Returns the *FIRST* part matching $name
     #>
    [OOXMLPart]getPartByName([string]$name){
        foreach($OOXMLPart in $this){
            if($OOXMLPart.getName() -eq $name){
                return $OOXMLPart
            }
        }
        return $null
    }

    <#
     # Returns the *FIRST* part matching $type
     #>
    [OOXMLPart]getPartByType([string]$type){
        foreach($OOXMLPart in $this){
            if($OOXMLPart.getType() -eq $type){
                return $OOXMLPart
            }
        }
        return $null
    }
}
