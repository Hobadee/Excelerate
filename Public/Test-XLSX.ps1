function Test-XLSX {
	<#
		.Synopsis
			Testing my Excel classes
		
		.Description
            Testing my Excel classes.

		.Parameter Filename
			Specify the Excel file to operate on
						
		.Example
			Test-XLSX -Filename <filename>
	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[parameter(Mandatory=$true)]
		[String]
		$Filename
	)

	$xlsx = [XLSX]::new($filename)
	return $xlsx

	<#
	echo "Status:  [$(New-Text "Success" -fg "Green")]"
	echo "Status:  [$(New-Text "Failure" -fg "Red")]"
	#>
}
