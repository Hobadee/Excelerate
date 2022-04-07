Using Module ../../../build/excelerate/excelerate.psd1


InModuleScope Excelerate{
	Describe "OOXML Class" {

        Context "Test loading blank OOXML files" {
            BeforeAll{
                $dir = "Tests"
            }
            BeforeEach{
                $OOXML = [OOXML]::new()
            }
            AfterEach{
                $OOXML.Dispose()
            }
            It "Checking DOCX"{
                $fname = "blank.docx"
                $file = (Join-Path $dir $fname)
                $OOXML.setFilename($file).load()
                $OOXML.checkOOXML() | Should -BeTrue
            }
            It "Checking PPTX"{
                $fname = "blank.pptx"
                $file = (Join-Path $dir $fname)
                $OOXML.setFilename($file).load()
                $OOXML.checkOOXML() | Should -BeTrue
            }
            It "Checking xlsx"{
                $fname = "blank.xlsx"
                $file = (Join-Path $dir $fname)
                $OOXML.setFilename($file).load()
                $OOXML.checkOOXML() | Should -BeTrue
            }
        }


        Context "Test loading fake OOXML files" {
            BeforeAll{
                $dir = "Tests"
            }
            BeforeEach{
                $OOXML = [OOXML]::new()
            }
            AfterEach{
                $OOXML.Dispose()
            }

            It "Checking Fake File"{
                $fname = "fakeEmpty.docx"
                $file = (Join-Path $dir $fname)
                {$OOXML.setFilename($file).load()} | Should -Throw
                $OOXML.checkOOXML() | Should -BeFalse
            }
            It "Checking Fake Zip File"{
                $fname = "fakeZip.docx"
                $file = (Join-Path $dir $fname)
                {$OOXML.setFilename($file).load()} | Should -Throw
                $OOXML.checkOOXML() | Should -BeFalse
            }
        }


        Context "Test saving" {
            BeforeAll{
                #$dir = "Tests"
            }
            AfterEach{
                #$OOXML.Dispose()
                #Remove-Item ($saveFile + ".*")
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

                #$saved = $XLSX.saveAs($saveFile)
                #$saved | Should -BeOfType System.IO.FileInfo
                #$saved.FullName | Should -Be $saveFile
            }
            It "Checking Save As overwrite behavior"{
                #{ $XLSX.saveAs($saveFile) } | Should -Throw
            }
        }
	}
}
