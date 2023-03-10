<#
.SYNOPSIS
Collection of [XLSXReference] objects
#>
class XLSXReferences : System.Collections.Generic.List[PSObject]{


    <#
    .SYNOPSIS
    Return a collection of references that match a name search query

    .DESCRIPTION
    This does a "-like" search to see if the "name" field is equal
    #>
    [XLSXReferences]getRefsByName([string]$search){

        $refs = [XLSXReferences]::new()

        foreach($item in $this){
            if($item.name -like $search){
                $refs.add($item)
            }
        }

        return $refs
    }


    <#
    .SYNOPSIS
    Return a collection of references that match a text search query

    .DESCRIPTION
    This does a "-like" search to see if the "#text" field is equal

    .EXAMPLE
    Search for broken refs:
    getRefsByText('*#REF!*')
    #>
    [XLSXReferences]getRefsByText([string]$search){

        $refs = [XLSXReferences]::new()

        foreach($item in $this){
            if($item.'#text' -like $search){
                $refs.add($item)
            }
        }

        return $refs
    }
}
