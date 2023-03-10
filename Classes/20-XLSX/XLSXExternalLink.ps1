<#

For External Links, read `rId` from `xl/workbook.xml` and cross reference to
`<Relationship Id="rId##">` in `xl/_rels/workbook.xml` to get the `Target`.

We load `Target` and get a bunch of (useless?) sheet data.  We load
`Target/_rels/Target.rels` and get a new `Target` (URI) which we can test for
connectivity.  If no connectivity, delete all references.

As part of cleanup, ensure that neither is orphaned and removed if orphaned.

#>


class XLSXExternalLink{

    <#
    Reference - The reference to the external link in the root workbook file.  This JUST contains an rId.
    Relationship - The relationship of the rId to the XML file containing the link.
    externalLink - The XML file describing the actual link
    #>
    [System.Xml.XmlElement] $reference
    [System.Xml.XmlElement] $relationship
    [XmlStream] $externalLink
    [XmlStream] $externalLinkRel

    [boolean]$UriTested = $false
    [boolean]$UriWorks = $null


    <#
    .SYNOPSIS
    Class contructor
    #>
    XLSXExternalLink([System.Xml.XmlElement]$reference, [XLSX]$xlsx){
        $this.reference = $reference

        <#
        To get target:
        1. load /xl/_rels/workbook.xml.rels
        2. match RID
        3. load "Target" attribute
        #>
        #$target = 

        # Search for the "rId" in the relation XML
        # and pull the relation data
        $index = $xlsx.relations.DocumentElement.Relationship.Id.IndexOf($this.getRID())
        $this.relationship = $xlsx.relations.DocumentElement.Relationship[$index]

        $fileLink = '/xl/'+$this.relationship.Target

        try{
            $this.externalLink = $xlsx.getXmlFile($fileLink)
        }
        catch{
            # External link doesn't exist and must be invalid.  Recommend for cleanup
            [Logging]::Debug("Failed to load '"+$fileLink+"' from XLSX")
            $this.externalLink = $null
            $this.forceURIStatus($false)
        }

        <#
        # At this point we have a list of all the RIDs.
        # We now need to load /xl/_rels/workbook.xml.rels
        # Then we tie each RID to a Target
        # Then we load each Target (in relation to the "/xl/" folder)
        # And we can load those as actual ExternalLink objects to check
        # The XLSXExternalLinks object should probably track the workbook.xml.rels and modify it as needed
        #>
    }


    <##
     # Tests if the links URI works or not.
     # Will only test once - subsequent runs will return cached value
     #>
    [boolean]testURI(){
        # Check if we have tested already and return cached result if so
        if($this.UriTested -eq $true){
            return $this.UriWorks
        }

        # Offload testing to the Test-Uri module I created
        $this.UriWorks = Test-Uri $this.targetURI
        $this.UriTested = $true

        return $this.UriWorks
    }


    <##
     # Invalidate the cached or manual test result and force a retest of the URI
     #
     # NOTE: This will override a manually set URI status!
     #
     # @return boolean $true if URI works, $false if not
     #>
    [boolean]forceTestURI(){
        $this.UriTested = $false
        return $this.testURI()
    }


    <##
     # Forces the URI check to a manual value
     # Useful for manually overriding URIs before later removal of non-working URIs
     #
     # $var boolean Status of the URI in question.  $true = working, $false = nonworking
     # @return XLSXExternalLink Returns this object
     #>
    [XLSXExternalLink]forceURIStatus([boolean]$status){
        $this.UriTested = $true
        $this.UriWorks = $status
        return $this
    }


    <##################
    ##### GETTERS #####
    ##################>
    [string]getRID(){
        return $this.reference.id
    }
    [string]getTargetURI(){
        return $this.targetURI
    }


}
