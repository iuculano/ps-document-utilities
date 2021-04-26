function Import-ExcelDocument
{
    <#
        .SYNOPSIS
        Imports an Excel document.
        
        .DESCRIPTION
        Imports an Excel document.

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
        $data = Import-ExcelDocument ".\SomeExcelFile.xlsx"

        Filtering:
        $data     = Import-ExcelDocument ".\SomeExcelFile.xlsx"
        $filtered = $data | Where-Object { $_.SomeColumn -like "Whatever" }


        .NOTES
        Does not require Excel.

        This function works by treating the document as a ZIP file, extracting
        its contents, then parsing the XML. This temporarily writes to $env:TEMP.


        Import-ExcelDocument
        Alex Iuculano, 2020
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
            # Quick sanity check for the file extension, the old format will not work
            $extension = $p.Path.Substring($p.Path.LastIndexOf(".") + 1)
            if ($extension -ne "xlsx")
            {
                Write-Error "Invalid file extension ($extension) - expected 'xlsx'."
                continue
            }


            # Create a unique name for the path, never want to conflict
            $handle   = New-Guid
            $tempPath = "$env:TEMP\ax\ExcelScratch\$handle"

            # Treat the Word document as a ZIP and extract it to the temp directory
            # New-Item pipes to the void because it'll barf the directory into the
            # pipeline otherwise, which will naturally return back...
            New-Item $tempPath -ItemType Directory -Force | Out-Null
            Copy-Item $p -Destination "$tempPath\Document.zip" -Force
            Expand-Archive "$tempPath\Document.zip" -DestinationPath $tempPath

            
            # Get ready to parse the XML
            $xml                    = [XML]::new()
            $xml.PreserveWhitespace = $true


            # First, load the sharedStrings.xml file
            # This is lookup table of text in the document
            $xml.Load("$tempPath\xl\sharedStrings.xml")
            $data = $xml.sst.si.t # Walk down the tags

            # Rows effectively holds the indicies to the actual data
            $xml.Load("$tempPath\xl\worksheets\sheet1.xml")
            $rows   = $xml.worksheet.sheetData.row
            $schema = $rows[0].c | ForEach-Object { $data[$_.v] }

            # $schema = # $data[0..$((@($rows.c).Count)-1)]


            $table = @()
            foreach ($row in $rows[1..$($rows.Length - 1)])
            {
                # Stage the object and "allocate" all the columns
                # The ordered type is paramount here
                $object = [Ordered]@{}
                for($i = 0; $i -lt $schema.Length; $i++)
                {
                    $object[$schema[$i]] = ""
                }


                # Build the lookup table
                foreach ($column in $row.c)
                {
                    # Decode column letters to index
                    if (!($column.r -match "[A-Z]{0,4}"))
                    {
                        throw "Failed to parse column."
                    }

                    # Will get the column number even from ones far down the line
                    # like 'AA', 'AB', etc.
                    $index        = 0
                    $columnString = $Matches[0]     
                    for ($i = 0; $i -lt $columnString.Length; $i++)
                    {
                        # Handle the "rollover" - since it's 0, first round is a no-op.
                    
                        # If there's another character, we've exhausted the character set...
                        # Multiply by the length of the character set each time we need to
                        # roll over like this and finally add what the character contributes
                        # ALX                       = 0x‭03F4‬
                        # (0x0000 * 0 ) + A (1 dec) = 0x0001 (1   dec)
                        # (0x0001 * 26) + L (5 dec) = 0x001F (31  dec)
                        # (0x001F * 26) + F (6 dec) = 0x032C (812 dec)
                        $index *= 26;
                    
                        # Convert the column character to a digit - remember, it's a 1-based index
                        # For example:
                        # A = 1, Z = 26
                        $index += ($columnString[$i] - [Byte][Char]"A") + 1
                    }
                    

                    # Try to discern text from numbers - checking for this
                    # tag seems to work well enough, in practice
                    $text = $column.Attributes.Where({ $_.Name -eq "t" })
                    if ($text.Count -gt 0)
                    {
                        $object[$index - 1] = $data[$column.v]
                    }

                    else
                    {
                        $object[$index - 1] = $column.v
                    }
                }
                
                $table += [PSCustomObject]$object
            }

            
            # Try to clean up and blow away the extracted data
            Remove-Item -LiteralPath $tempPath -Recurse -Force

            
            # Output
            $table
        }
    } # Process
}
