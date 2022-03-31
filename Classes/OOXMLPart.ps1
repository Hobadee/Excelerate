class OOXMLPart{
    [string]$name
    [string]$type


    OOXMLPart([string]$name, [string]$type){
        $this.name = $name
        $this.type = $type
    }


    [string]getName(){
        return $this.name
    }
    [string]getType(){
        return $this.type
    }
}
