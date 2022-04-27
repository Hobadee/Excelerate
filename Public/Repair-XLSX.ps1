function Repair-XLSX {
	<#
		.Synopsis
			Repairs a slow-loaing Excel file
		
		.Description
            Repairs slow-loading Excel files by removing broken references and other items from the file.

		.Parameter InFile
			Specify the Excel file to operate on

		.Parameter OutFile
			Specify the filename to save the fixed Excel file as
			NOTE: This is not used if the Overwrite flag is used.
		
		.Parameter Overwrite
            Overwrite the original file instead of making a copy
            NOTE: This is dangerous, especially with the -Force option!  You are entering the data-loss zone!
		
		.Parameter Force
			Force all operations - don't prompt for confirmation.

		.Parameter CleanDefinedName
			Clean Defined Names from the XLSX file
			This will remove all Defined Names which are either hidden, or have invalid references

		.Parameter CleanExternalReference
			Clean up External References
			This will check all External References and delete them if they aren't accessible.

		.Parameter CleanAll
			Clean all items that are able to be cleaned.  Good for fully-automated use cases
				
		.Example
			Repair-XLSX -Filename <filename>
			Check <filename> for issues and ask to repair them
	#>
	[CmdletBinding(SupportsShouldProcess=$true)]
	param (
		[parameter(Mandatory=$true)]
		[String]
		$Infile
	,
		[String]
		$Outfile
	,
		[Switch]
		$Overwrite
	,
		[Switch]
		$Force
	,
		[Switch]
		$CleanDefinedName
	,
		[Switch]
		$CleanExternalReference
	,
		[Switch]
		$CleanAll
	)
	<#
		CmdletBinding/ShouldProcess parameters we should specifically be aware of/handle:
		WhatIf, Confirm, Verbose, Debug


		Use cases:
		-User wants to do an automated fix in the background, saving to a new file
			Flags: -Force -
		-User wants to do an automated fix in the background, overwriting the existing file
		-User wants to do an interactive fix, saving to a new file
		-User wants to do an interactive fix, overwriting the existing file
		-User wants a report of how many errant items are detected
	#>

	return $false

	<#
	echo "Status:  [$(New-Text "Success" -fg "Green")]"
	echo "Status:  [$(New-Text "Failure" -fg "Red")]"
	#>
}
