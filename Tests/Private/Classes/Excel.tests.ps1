#Using Module ../../../build/excelerate/excelerate.psd1


InModuleScope Excelerate{
	Describe "Excel Class" {

        Context "Testing blank Excel object" {
            BeforeAll{
                $xlDir = "Tests"
                $fname = "blank.xlsx"
                $workbookFile = "xl/workbook.xml"
                $relsFile = "_rels/.rels"
                $saveFile = "/Users/eric.kincl/src/excelerate/testExcel.xlsx"
                $xl = (Join-Path $xlDir $fname)
                $excel = [Excel]::new($xl)
            }
            AfterAll{
                $excel.Dispose()
                Remove-Item $saveFile
            }


            It "Checking object type"{
                # Pester assertion "BeOfType" doesn't work with custom types
                $excel.getType().Name | Should -Be "Excel"
            }
            It "Checking files exist"{
                $excel.getTempDir() | Should -exist
                (Join-Path $excel.getTempDir() $workbookFile ) | Should -exist
                (Join-Path $excel.getTempDir() $relsFile ) | Should -exist
            }


            It "Checking Getter Types"{
                $excel.getTempDir() | Should -BeOfType string
                $excel.getWorkbook() | Should -BeOfType xml
                $excel.getFilename() | Should -BeOfType string
            }
            It "Checking Getters"{
                $excel.getFilename() | Should -Be $fname
            }

            It "Checking Save As operations"{
                # Test that "saveAs" without a dir saves to current directory
                # Test that "saveAs" with a dir saves to the correct directory
                # Test that saving on an existing file does/doesn't overwrite it

                $saved = $excel.saveAs($saveFile)
                $saved | Should -BeOfType System.IO.FileInfo
                $saved.FullName | Should -Be $saveFile
            }
            It "Checking Save As overwrite behavior"{
                { $excel.saveAs($saveFile) } | Should -Throw
            }
        }		
	}
}
