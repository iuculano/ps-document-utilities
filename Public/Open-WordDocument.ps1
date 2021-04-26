function Open-WordDocument
{
    <#
        .SYNOPSIS
        Opens a Word document for manipulation.
        
        .PARAMETER LiteralPath
        The path of the input file that will be converted - used as is.
        
        Wildcards are NOT functional with this parameter, as it expects a
        true, direct path to the location.

        .PARAMETER Path
        The path of the input file that will be converted. Supports wildcards.

        .PARAMETER ReadOnly
        Opens the file as ReadOnly. No modificaitons will be saved.

        .PARAMETER Hidden
        Determines whether the underlying Word instance is hidden.
        
        Word is visible by default just because it's so easy to leave a
        hanging instance.

        .EXAMPLE
        Opening a file:
        $document       = Open-WordDocument "Some file.docx"

        Replacing and bolding a certain word:
        $document       = Open-WordDocument "Some file.docx"
        $document.Words | ForEach-Object { $_.Bold = $false }
        $document.Words | ForEach-Object {

            # Note the space after!
            if ($_.Text -eq "REPLACE_ME ")
            {
                $_.Text = "AWESOME_NEW_STRING "
                $_.Bold = $true
            }
        }

        .NOTES
        Requires Word.

        This also has a nasty side effect that you can leave a hanging Word instance.
        Always call Close-WordDocument when you're done!

        Be very careful that you're actually accessing the reference to the
        object you intend to edit and watch for testing against text.

        Word stores words with the trailing whitespace because I have no idea why.
        Whitespace is not actually the delimiter, rather the next character after
        whitespace.

        Every day we drift further from God's light.


        Open-WordDocument.ps1    
        Alex Iuculano, 2019
    #>

    [CmdletBinding(DefaultParameterSetName = "Path")]
    Param
    (
        [Alias("PSPath")]
        [Parameter(ParameterSetName                = "LiteralPath",
                   ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [ValidateNotNullOrEmpty()]
        [String[]]$LiteralPath,

        [Parameter(Position                        = 0,
                   ParameterSetName                = "Path",
                   ValueFromPipeline               = $true,
                   ValueFromPipelineByPropertyName = $true)]
        [ValidateScript({ Test-Path $_ })]
        [ValidateNotNullOrEmpty()]
        [String[]]$Path,

        [Switch]$ReadOnly = $false,
        [Switch]$Hidden   = $false
    )


    Begin
    {
        # This feels yucky but Close-WordDocument can't function correctly unless
        # you assign a Word instance per document. 
        
        # Otherwise, if you made 2 seperate calls to Open-WordDocument, it would
        # be impossible to know which application to close. It's annoyingly easy
        # to leave yourself with a hanging instance - consider yourself warned.

        # Because of this, this function creates a visible Word instance by default.
    }

    Process
    {
        $PathList                             = @{ }
        $PathList[$PSCmdlet.ParameterSetName] = (Get-Variable $PSCmdlet.ParameterSetName).Value
        $PathList                             = Resolve-Path @PathList

        foreach ($p in $PathList)
        {
            # Mainly for information about the files
            $files = Get-ChildItem -LiteralPath $p         
            
            
            foreach ($file in $files)
            {
                try 
                {
                    # https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
                    Write-Verbose "Preparing Word COM interface."                    
                    $word         = New-Object -ComObject Word.Application
                    $word.Visible = -not [Bool]$Hidden


                    # https://docs.microsoft.com/en-us/office/vba/api/word.documents.open
                    Write-Verbose "Opening file -> $($file.FullName)"
                    $document = $word.Documents.Open($file.FullName, [Bool]$ReadOnly, -not [Bool]$Hidden)
                    if ([Bool]$document.Application)
                    {
                        # Save some typing
                        $content = $document.Content


                        # Semi-abstracted and simplified object

                        # Note that this is very much LIVE data and any change
                        # you make will immediately reflect in the document itself!
                        [PSCustomObject]@{

                            # Underlying type name - if you edit this for whatever
                            # reason, make sure you update the Param checks for it!
                            PSTypeName = "ax.WordDocument"


                            # The low level COM objects representing the document
                            # The Word instnace MUST stay alive for this to remain valid
                            COMObjectWord     = $word
                            COMObjectDocument = $document


                            # Info about the underlying file being read
                            File = $file

                            
                            # Whether we're opened in ReadOnly mode
                            IsReadOnly = $ReadOnly

                            # Whether we're opened in Hidden mode
                            IsHidden   = $Hidden


                            # Office uses 1-based indexing because Microsoft 
                            # programmers are apparently complete animals. 
                            
                            # Wrapping these like this has the nice effect of 
                            # yielding a 0-indexed array - and since PowerShell
                            # passes objects by reference, everything *just works*.
                            Start      = $content.Start
                            End        = $content.End

                            Text       = $content.Text
                            Characters = @($content.Characters)
                            Words      = @($content.Words)
                            Sentences  = @($content.Sentences)
                            Paragraphs = @($content.Paragraphs)

                            Tables     = @($content.Tables)
                            Footnotes  = @($content.Footnotes)
                            Endnotes   = @($content.Endnotes)
                            Comments   = @($content.Comments)
                            Sections   = @($content.Sections)
                            Bookmarks  = @($content.Bookmarks)
                        }
                    }
                }

                catch [Runtime.InteropServices.COMException]
                {
                    # Going down in flames, try to clean up...
                    $word.Application.Quit()


                    $exception = $_        
                    switch ($_.Exception.HResult) 
                    {
                        # Can't find the COM interface
                        0x80040154
                        {
                            throw "Word COM Interface not found - is Office installed?"
                        }

                        # File error
                        0x800A03EC
                        {
                            # Annoyingly, the HRESULT is the same for 
                            # wrong format / file and a completely bad path...
                            Write-Error "Unrecognized or invalid file format."
                        }
        
                        # Grab bag
                        default
                        {
                            $PSCmdlet.ThrowTerminatingError($exception)
                        }
                    } # switch
                } # catch

                catch
                {
                    # Going down in flames, try to clean up...
                    $word.Application.Quit()
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            } # foreach
        }
    } # Process
}
