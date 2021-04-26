function Close-WordDocument
{
    <#
        .SYNOPSIS
        Closes an open Word Document.
        
        .PARAMETER WordDocument
        Specifies an object representing an open Word Document.

        .NOTES
        Requires Word.

        Does what it says on the tin. Clean up after yourself.


        Close-WordDocument.ps1    
        Alex Iuculano, 2019
    #>

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory                       = $true, 
                   ValueFromPipeline               = $true,
                   ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ $_.PSObject.TypeNames[0] -eq "ax.WordDocument" })]
        [PSCustomObject[]]$WordDocument
    )


    Process
    {
        foreach ($doc in $WordDocument)
        {
            try
            {
                # https://docs.microsoft.com/en-us/office/vba/api/word.document.close(method)
                # Never save changes (not that there are any), this is always a completely ReadOnly operation

                # If you need to save, use the Save-WordDocument function
                [Void]$doc.COMObjectDocument.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
                [Void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc.COMObjectDocument)
                     

                [Void]$doc.COMObjectWord.Application.Quit()
                [Void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc.COMObjectWord)    
            }

            catch
            {
                $PSCmdlet.ThrowTerminatingError($_)
            }


            $doc.COMObjectDocument = $null
            $doc.COMObjectWord     = $null
        }
    } # Process
}
