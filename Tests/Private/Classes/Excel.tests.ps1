Using Module ../../../build/excelerate/excelerate.psd1


InModuleScope Excelerate{
	Describe "XLSX Class" {

        Context "Testing blank XLSX object" {
            BeforeAll{
                $xlDir = "Tests"
                $fname = "blank.xlsx"
                $workbookFile = "xl/workbook.xml"
                $relsFile = "_rels/.rels"
                $saveFile = "/Users/eric.kincl/src/excelerate/testExcel.xlsx"
                $xl = (Join-Path $xlDir $fname)
                $XLSX = [XLSX]::new($xl)
            }
            AfterAll{
                $XLSX.Dispose()
                Remove-Item $saveFile
            }


            It "Checking object"{
                # Pester assertion "BeOfType" doesn't work with custom types
                $XLSX.getType().Name | Should -Be "XLSX"
                $XLSX.checkXLSX() | Should -BeTrue
            }
            It "Checking files exist"{
                # This should be moved to the OOXML test
                <#
                $XLSX.getTempDir() | Should -exist
                (Join-Path $XLSX.getTempDir() $workbookFile ) | Should -exist
                (Join-Path $XLSX.getTempDir() $relsFile ) | Should -exist
                #>
            }

            It "Checking Save As operations"{
                # Test that "saveAs" without a dir saves to current directory
                # Test that "saveAs" with a dir saves to the correct directory
                # Test that saving on an existing file does/doesn't overwrite it

                $saved = $XLSX.saveAs($saveFile)
                $saved | Should -BeOfType System.IO.FileInfo
                $saved.FullName | Should -Be $saveFile
            }
            It "Checking Save As overwrite behavior"{
                { $XLSX.saveAs($saveFile) } | Should -Throw
            }
        }		
	}
}
