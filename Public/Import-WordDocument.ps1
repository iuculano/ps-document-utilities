function Import-WordDocument
{
    <#
        .SYNOPSIS
        Imports a Word document.
        
        .DESCRIPTION
        Imports a Word document.

        This returns a plain text version of the document, with no particular
        metadata for extended information. The parser is very barebones and
        will not include any formatting beyond basic text structure.

        .PARAMETER LiteralPath
        Specifies the path of the input file - used as is.
        
        Wildcards are NOT functional with this parameter, as it expects a
        true, direct path to the location.

        .PARAMETER Path
        Specifies the path of the input file - supports wildcards.

        .EXAMPLE
        Opening a file:
        $text = Open-WordDocument "Some file.docx"

        Regex search:
        $text = Open-WordDocument "Some file.docx"
        switch -regex ($text)
        {
            "{SOME_STRING}"
            {
                Do-Thing
            }
        }


        .NOTES
        Does not require Word.

        This function works by treating the document as a ZIP file, extracting
        its contents, then parsing the XML. This temporarily writes to $env:TEMP.

        See here for some more information about the query:
        https://www.w3schools.com/xml/xpath_syntax.asp
        https://stackoverflow.com/questions/35606708/what-is-the-difference-between-and-in-xpath/35606964

        Import-WordDocument
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
        [String[]]$Path
    )


    Process
    {
        $pathList                             = @{ }
        $pathList[$PSCmdlet.ParameterSetName] = (Get-Variable $PSCmdlet.ParameterSetName).Value
        $pathList                             = Resolve-Path @pathList

        foreach ($p in $pathList)
        {
            # Create a unique name for the path, never want to conflict
            $handle   = New-Guid
            $tempPath = "$env:TEMP\ax\WordScratch\$handle"

            # Treat the Word document as a ZIP and extract it to the temp directory
            # New-Item pipes to the void because it'll barf the directory into the
            # pipeline otherwise, which will naturally return back...
            New-Item $tempPath -ItemType Directory -Force | Out-Null
            Copy-Item $p -Destination "$tempPath\Document.zip" -Force
            Expand-Archive "$tempPath\Document.zip" -DestinationPath $tempPath

            
            # Get ready to parse the XML
            $xml                    = [XML]::new()
            $xml.PreserveWhitespace = $true
            $xml.Load("$tempPath\word\document.xml")

            # All the tags we need will be prefixed with 'w'
            $nt  = [System.Xml.NameTable]::new()
            $nsm = [System.Xml.XmlNamespaceManager]::new($nt)
            $nsm.AddNamespace("w", $xml.document.w)


            $buffer         = [Text.StringBuilder]::new()
            $paragraphNodes = $xml.document.ChildNodes.SelectNodes("//w:p", $nsm)
            foreach ($paragraph in $paragraphNodes)
            {
                $textNodes = $paragraph.SelectNodes(".//w:t | .//w:tab | .//w:br", $nsm)
                foreach ($text in $textNodes)
                {
                    switch ($text.Name)
                    {
                        "w:t"
                        {
                            [Void]$buffer.Append($text.InnerText)
                        }

                        "w:tab"
                        {
                            [Void]$buffer.Append("`t")
                        }

                        "w:br"
                        {
                            [Void]$buffer.Append("`n")
                        }
                    }
                }


                # Not 100% sure this is technically correct (paragraphs innately line-breaking)
                # but it definitely looks considerably better...
                [Void]$buffer.Append("`n")
            }

            
            # Try to clean up and blow away the extracted data
            Remove-Item -LiteralPath $tempPath -Recurse -Force


            # Output
            $buffer.ToString().Trim()
        } #foreach
    } # Process
}
