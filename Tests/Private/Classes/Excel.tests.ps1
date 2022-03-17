#Using Module ../../../build/excelerate/excelerate.psd1


InModuleScope Excelerate{
	Describe "Excel Class" {

        Context "Testing blank Excel object" {
            BeforeAll{
                $xlDir = "Tests"
                $fname = "blank.xlsx"
                $workbookFile = "xl/workbook.xml"
                $xl = (Join-Path $xlDir $fname)
                $excel = [Excel]::new($xl)
            }
            AfterAll{
                $excel.Dispose()
            }


            It "Checking object type"{
                # Pester assertion "BeOfType" doesn't work with custom types
                $excel.getType().Name | Should -Be "Excel"
            }
            It "Checking files exist"{
                $excel.getTempDir() | Should -exist
                (Join-Path $excel.getTempDir() $workbookFile ) | Should -exist
            }

            
            It "Checking Getter Types"{
                $excel.getTempDir() | Should -BeOfType string
                $excel.getWorkbook() | Should -BeOfType xml
                $excel.getFilename() | Should -BeOfType string
            }
            It "Checking Getters"{
                $excel.getFilename() | Should -Be $fname
            }
        }		
	}
}
