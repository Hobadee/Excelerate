function Repair-XLSX {
	<#
		.Synopsis
			Repairs a slow-loaing Excel file
		
		.Description
            Repairs slow-loading Excel files by removing broken references and other items from the file.

		.Parameter Filename
			Specify the Excel file to operate on
		
		.Parameter Overwrite
            Overwrite the original file instead of making a copy
            NOTE: This is dangerous, especially with the -Force option!  You are entering the data-loss zone!

		.Parameter CleanDefinedNames
			Clean Defined Names from the XLSX file
			This will remove all Defined Names which are either hidden, or have invalid references

		.Parameter CleanExternalReferences
			Clean up External References
			This will check all External References and delete them if they aren't accessible.
				
		.Example
			Repair-XLSX -Filename <filename>
			Check <filename> for issues and ask to repair them
	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[parameter(Mandatory=$true)]
		[String]
		$Filename
	,
		[Switch]
		$Overwrite
	,
		[Switch]
		$CleanDefinedNames
	,
		[Switch]
		$CleanExternalReferences
	)

	return $false

	<#
	echo "Status:  [$(New-Text "Success" -fg "Green")]"
	echo "Status:  [$(New-Text "Failure" -fg "Red")]"
	#>
}
