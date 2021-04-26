. "$PSScriptRoot\..\..\Public\Import-WordDocument.ps1"


Describe "Import-WordDocument" {

    BeforeAll {

        # Note that if you add more Test documents, they'll be picked up here
        $TestDocuments = Get-ChildItem "$PSScriptRoot\..\Data\Test_*.docx"

        # Ensure that this ALWAYS matches the FIRST document
        $TestText      = "Testing, testing - 1, 2, 3!`nSecond line."
    }


    Context "Parameter validation" {

        It "LiteralPath/Path - throw on bad or invalid path." {

            $testPaths = 
            @(
                "",
                "THIS*SHOULD*NEVER*MATCH",
                "BAD_PATH",
                @(),
                $null
            )

            foreach ($path in $testPaths)
            {
                { Import-WordDocument -LiteralPath $path } | Should -Throw
                { Import-WordDocument -Path $path        } | Should -Throw
            }
        }

        It "LiteralPath/Path - return data from valid path" {

            # Load 1 at a time
            foreach ($path in $TestDocuments)
            {
                { Import-WordDocument -LiteralPath $path.FullName  } | Should -BeTrue
                { Import-WordDocument -Path $path.FullName         } | Should -BeTrue
            }


            # Load an array
            $TestDocuments | Import-WordDocument                                  | Should -BeTrue
            { Import-WordDocument -LiteralPath $TestDocuments                       } | Should -BeTrue            
            { Import-WordDocument -Path "$($TestDocuments[0].DirectoryName)\*.docx" } | Should -BeTrue              
        }
    }
    
    Context "Data validation" {

        BeforeAll {

            $docs = Import-WordDocument -LiteralPath $TestDocuments
            $docs | Should -BeTrue
        }

        It "Should return $($TestDocuments.Count) documents" {

            $docs | Should -HaveCount $TestDocuments.Count    
        }

        It "Should return objects of type 'String'" {

            foreach ($d in $docs)
            {
                $d.GetType().Name | Should -Be "String"
            }
        }

        It "Should read the file correctly" {

            $docs[0] | Should -Be $TestText
        }

        It "Should be no leading or trailing whitespace" {

            [Char]::IsWhiteSpace([Char]$docs[0][0] ) | Should -BeFalse
            [Char]::IsWhiteSpace([Char]$docs[0][-1]) | Should -BeFalse
        }
    }
}
