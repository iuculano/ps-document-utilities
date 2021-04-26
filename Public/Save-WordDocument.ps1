function Save-WordDocument
{
    <#
        .SYNOPSIS
        Saves a Word document.

        .DESCRIPTION
        Saves a Word document.
        
        potentially to a different (Word supported) format.

        .PARAMETER InputObject
        Word Document object being operated on.
        
        .PARAMETER LiteralPath
        The path of the input file that will be converted - used as is.
        
        Wildcards are NOT functional with this parameter, as it expects a
        true, direct path to the location.

        .PARAMETER Path
        The path of the input file that will be converted. Supports wildcards.

        .PARAMETER FileType
        The file type to convert to.
        By default, this will assume PDF.
        
        .NOTES
        Requires Word.

        Relative paths are supported.


        Save-WordDocument.ps1    
        Alex Iuculano, 2018
    #>

    [CmdletBinding(DefaultParameterSetName = "Path")]
    Param
    (
        [Parameter(Position                        = 0,
                   Mandatory                       = $true, 
                   ValueFromPipeline               = $true,
                   ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ $_.PSObject.TypeNames[0] -eq "ax.WordDocument" })]
        [PSCustomObject]$InputObject,

        [Alias("PSPath")]
        [Parameter(ParameterSetName                = "LiteralPath",
                   ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ Test-Path $_ -IsValid })]
        [String]$LiteralPath,

        [Parameter(Position                        = 1,
                   ParameterSetName                = "Path",
                   ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ Test-Path $_ -IsValid })]
        [String]$Path,

        [ValidateSet("Default", "Word", "Legacy", "PDF", "Html", "PlainText")]
        [String]$FileType = "PDF"
    )


    $fileExtension  = "this.is.a.bug.if.you.ever.see.this.tell.alex"
    $fileFormatEnum = switch ($FileType)
    {
        @("Default", "Word")
        {
            $fileExtension = "docx"
            [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
        }

        "Legacy"
        {
            $fileExtension = "doc"
            [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocument
        }

        "PDF"
        {
            $fileExtension = "pdf"
            [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatPDF
        }

        "Html"
        {
            $fileExtension = "html"
            [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatHTML
        }

        "PlainText"
        {
            $fileExtension = "txt"
            [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatTextLineBreaks
        }

        default
        {
            throw "Invalid file type - $FileType."
        }
    }


    $outputPath = (Get-Variable $PSCmdlet.ParameterSetName).Value
    if (!$outputPath)
    {
        $outputPath = "$((Get-Location).Path)\$($InputObject.File.BaseName).$fileExtension"
    }
    

    # If $outputPath is relative, resolve to an absolute path
    $outputPath = [IO.Path]::GetFullPath("$outputPath")


    # Annoyingly, the SaveAs() method seems to fail completely and utterly silently...
    # It simply does nothing for a bad path, so there's nothing to catch or test
    $InputObject.COMObjectDocument.SaveAs($outputPath, $fileFormatEnum)

    # Need to check wheter the file exists at this point
    if (Test-Path $outputPath)
    {
        # Return whatever we wrote
        Get-ChildItem $outputPath
    }

    else
    {
        throw "Failed to write file '$outputPath'."
    }
}
