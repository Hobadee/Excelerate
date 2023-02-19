class XLSXDefinedName{

    [string]$name
    [string]$value
    [string]$localSheetId
    [boolean]$hidden = $false


    <##################
    ##### GETTERS #####
    ##################>
    [string]getName(){
        return $this.name
    }
    [string]getValue(){
        return $this.value
    }
    [string]getLocalSheetId(){
        return $this.localSheetId
    }
    [boolean]getHidden(){
        return $this.hidden
    }


    <##################
    ##### SETTERS #####
    ##################>
    [XLSXDefinedName]setName([string]$name){
        $this.name = $name
        return $this
    }
    [XLSXDefinedName]setValue([string]$value){
        $this.value = $value
        return $this
    }
    [XLSXDefinedName]setLocalSheetId([string]$Id){
        $this.localSheetId = $Id
        return $this
    }
    [XLSXDefinedName]setHidden([boolean]$hidden){
        $this.hidden = $hidden
        return $this
    }

}
