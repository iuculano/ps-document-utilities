. "$PSScriptRoot\..\..\Public\Save-WordDocument.ps1"


Describe "Save-WordDocument" {

    Context "Parameter validation" {



        Mock Open-WordDocument { 

            return 
        }

        BeforeAll {

            $badObject  = [PSCustomObject]@{ PSTypeName = "ax.NotAWordDocument" }
            $goodObject = [PSCustomObject]@{ PSTypeName = "ax.WordDocument"     }
        }


        It "InputObject - throw on invalid type" {

            { Save-WordDocument $badObject } | Should -Throw     
        }

        It "InputObject - allow valid type" {

            { Save-WordDocument $goodObject } | Should -Not -Throw     
        }

        It "LiteralPath - throw on bad path" {

            { Save-WordDocument $goodObject -LiteralPath "BAD_PATH"  } | Should -Throw 
            { Save-WordDocument $goodObject -LiteralPath "BAD_PATH*" } | Should -Throw
            { Save-WordDocument $goodObject -LiteralPath "BAD:\"     } | Should -Throw
        }
    }

    Context "Results" {

        It "Should return data for a valid lookup" {

            # $result = Get-HIBPPassword "test"
            # $result |  Should -BeTrue
        }

        It "Should return nothing for invalid lookup" {

            # $result = Get-HIBPPassword "12309FAKE_PASSWORD_THIS_WILL_NEVER_WORK12309"
            #$result | Should -BeFalse
        }
    }
}
