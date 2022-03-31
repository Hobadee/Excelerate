function Test-OOXML {
	<#
		.Synopsis
			Testing my OOXML classes
		
		.Description
            Testing my OOXML classes.

		.Parameter Filename
			Specify the OOXML file to operate on
						
		.Example
			Test-OOXML -Filename <filename>
	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[parameter(Mandatory=$true)]
		[String]
		$Filename
	)

	$ooxml = [OOXML]::new($filename)
	return $ooxml

	<#
	echo "Status:  [$(New-Text "Success" -fg "Green")]"
	echo "Status:  [$(New-Text "Failure" -fg "Red")]"
	#>
}
