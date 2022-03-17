Using Module ../../../build/excelerate/excelerate.psd1


InModuleScope Excelerate{
	Describe "tempDir Class" {

        Context "Testing Temp Dirs" {
            BeforeAll{
                $tmpDir = [tempDir]::new()
            }
            AfterAll{
                $tmpdir.Dispose()
            }


            It "getPath"{
                $tmpDir.getPath() | Should -BeOfType string
            }
            It "getName"{
                $tmpDir.getName() | Should -BeOfType string
            }
            It "getParent"{
                $tmpDir.getParent() | Should -BeOfType string
            }

            
            It "Does directory actually exist?"{
                $tmpDir.getParent() | Should -Exist
                $tmpDir.getPath() | Should -Exist
            }
            It "Is the temp directory removed?"{
                $tmpDir.Dispose()
                $tmpDir.getPath() | Should -Not -Exist
            }
        }		
	}
}
