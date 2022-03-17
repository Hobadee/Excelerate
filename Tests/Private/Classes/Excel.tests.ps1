#Using Module ../../../build/excelerate/excelerate.psd1


InModuleScope Excelerate{
	Describe "Excel Class" {

        Context "Testing blank Excel object" {
            BeforeAll{
                $xlDir = "Tests"
                $fname = "blank.xlsx"
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

            It "Checking Filename Type"{
                $excel.getFilename() | Should -BeOfType string
            }
            It "Checking Filename Value"{
                $excel.getFilename() | Should -Be $fname
            }
        }		
	}
}
