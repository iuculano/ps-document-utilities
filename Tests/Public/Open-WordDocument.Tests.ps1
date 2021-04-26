. "$PSScriptRoot\..\..\Public\Open-WordDocument.ps1"


Describe "Open-WordDocument" {

    BeforeAll {

        # Note that if you add more Test documents, they'll be picked up here
        $TestDocuments = Get-ChildItem "$PSScriptRoot\..\Data\Test_*.docx"

        # Ensure that this ALWAYS matches the FIRST document
        $TestText      = @("Testing, testing - 1, 2, 3!", "Second line.")
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
                { Open-WordDocument -LiteralPath $path } | Should -Throw
                { Open-WordDocument -Path $path        } | Should -Throw
            }
        }

        It "LiteralPath/Path - return data from valid path" {

            # Load 1 at a time
            foreach ($path in $TestDocuments)
            {
                { Open-WordDocument -LiteralPath $path.FullName  } | Should -BeTrue
                { Open-WordDocument -Path $path.FullName         } | Should -BeTrue
            }


            # Load an array
            { Open-WordDocument -LiteralPath $TestDocuments           } | Should -BeTrue            
            { Open-WordDocument -Path "$($path.DirectoryName)\*.docx" } | Should -BeTrue              
        }
    }
    
    Context "Data validation" {

        BeforeAll {

            $docs = Open-WordDocument -LiteralPath $TestDocuments
            $docs | Should -BeTrue
        }

        It "Should return $($TestDocuments.Count) documents" {

            $docs | Should -HaveCount $TestDocuments.Count    
        }

        It "Should return objects of type 'ax.WordDocument'" {

            foreach ($d in $docs)
            {
                $d.PSObject.TypeNames[0] | Should -Be "ax.WordDocument"
            }
        }

        It "Should have File bound" {

            $docs.File | Should -BeTrue
        }

        It "Should have read the file correctly (validate text)" {

            $docs[0].Sentences[0].Text | Should -BeLike "$($TestText[0])*"
            $docs[0].Sentences[1].Text | Should -BeLike "$($TestText[1])*"
        }


        # These tests aren't particularly good, but at least shows the data changed...
        It "Should respect Hidden flag" {

            $docs[0].IsHidden | Should -BeFalse
        }

        It "Should respect ReadOnly flag" {

            $docs[0].IsReadOnly | Should -BeFalse
        }
    }
}
