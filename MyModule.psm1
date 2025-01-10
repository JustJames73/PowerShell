<#
about_Functions				https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions
about_Functions_Advanced		https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced
about_Functions_Advanced_Methods	https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_methods
about_Functions_Advanced_Parameters	https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_advanced_parameters
about_Functions_CmdletBindingAttribute	https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_cmdletbindingattribute
about_Functions_OutputTypeAttribute	https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_functions_outputtypeattribute
#>

FUNCTION Search-Scripts {
<#
.SYNOPSIS
  Search all .PS1 files under the present working directory for keywords. 
.DESCRIPTION
  The Search-Scripts function combines the Get-ChildItem and Select-String
cmdlets to perform a keyword search on text files. By default the function 
recursively searches PowerShell script files (*.ps1) from the current folder 
using the provided keyword as a "simplematch". A simplematch searches the 
text as as string and does not evaluate Regular Expressions, if a RegEx is 
provided the search will return any string matching the keword as provided. 
  The path and file type can be changed with parameter switches. 
  The default view groups results by the file in which they were found but 
truncates the matching line of text, using the "-ListView" switch will show
the whole line of text matching the search. 
The Path is sent to "Out-GridView -Multiple", allowing selection of one or 
more scripts to open in PowerShellISE for further review. 
NOTE: the GridView window needs to close before you can use ISE. 
.EXAMPLE
	C:\> Search-Scripts -Keyword foreach
... will search all *.ps1 files from the root of C:\ containing the word "foreach"
.PARAMETER <Path>
    The path from where you wish to start your search. 
    Uses $pwd by default. 
.PARAMETER <Include>
    The filename pattern to include in the search.
    Uses "*.ps1" by defualt, but can be changed to any valid filename pattern. 
    Ex: "Scheduled*.ps1" or "*log.txt"
.PARAMETER <Keyword>
    The keword to search for. 
    This is a required parameter, and will prompt for a value if not provided. 
.PARAMETER <List>
    [Switch] to change the output from a formatted table grouping results by
    the path and filename with the matched line on a single truncated line 
    or displaying as a formatted list with the Path, Line Number, and full Line 
    of text matched by the search. 
    It is recomended that you narrow down the scope of your search with the 
    Path and Include parameters when using the ListView switch. 
 #>
    PARAM(
        [STRING[]]$Path = $pwd,
        [STRING[]]$Include = "*.ps1",
        [STRING[]]$KeyWord = (Read-Host "Keyword?"),
        [SWITCH]$ListView
    )
    BEGIN {

    }
    PROCESS {
        Get-ChildItem -path $Path -Include $Include -Recurse | `
	    sort Directory,CreationTime | `
	    Select-String -simplematch $KeyWord -OutVariable Result | `
        Out-Null
    }
    END {
        IF ($ListView) { 
            $Result | Format-List -Property Path,LineNumber,Line 
        } 
        ELSE { 
            $Result | Format-Table -GroupBy Path -Property LineNumber,Line -AutoSize 
        } 
        $Result | 
        select Path | sort Path -Unique | 
        Out-GridView -OutputMode Multiple -Title 'Select one or more files to open...' | 
        ForEach-Object { $psISE.CurrentPowerShellTab.Files.Add($_.Path) }  
    }
}
# https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_powershell_ise_exe?view=powershell-5.1
# https://learn.microsoft.com/en-us/powershell/scripting/windows-powershell/ise/introducing-the-windows-powershell-ise?view=powershell-7.3
# export the PATH to <Out-GridView -OutputMode Multiple> and then pass to ISE.exe to open in tabs
################################################################################

FUNCTION Add-TimeStamp
{
<#
.Synopsis
   Add a sortable timestamp to the filename using the last modified date of the file
.DESCRIPTION
   By default this function will replace file name with a sortable timestamp prefix
   in the directory where the file is currently located. 
   Use the -Suffix switch to add the timestamp at the end of the filename. 
   Use the -Copy switch to create a copy of the file with the new filename.
.EXAMPLE
   > Add-TimeStamp -FullName .\Foo.bar
   Renames the file "Foo.bar" to "2016-11-07T144837_Foo.bar"
.EXAMPLE
   > Add-TimeStamp .\Foo.bar -Suffix -CopyItem
   Copies the file "Foo.bar" to "Foo_2016-11-07T144837.bar"
.EXAMPLE
   > Get-ChildItem Foo*.bar | Add-TimeStamp
   Renames all files in the folder matching "Foo*.bar" with the last modified 
   timestamp. 
.NOTES
   This script was created to aid in creating backup copies of scripts and data files. 
#>
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  #HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [OutputType([String])]
  Param(
        # Param1 help description
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true,
                   HelpMessage="Enter the filename to prefix with a sortable timestamp. Use the Switch -Suffix to add the timestamp to the end of the filename",
                   Position=0)]
    [string] $FullName,
    [switch] $Suffix,
    [switch] $CopyItem
  )
      Begin
    {
    }
    Process
    {
        Write-Verbose -Message $("OldName:`t$((Get-ItemProperty $FullName).FullName)")
        $Vars = Get-ItemProperty $FullName 
        $Directory = $Vars.DirectoryName
        $BaseName = $Vars.BaseName
        $Extension = $Vars.Extension
        $TimeStamp = $Vars.LastWriteTime | Get-Date -Format yyy-MM-ddTHHmmss
        IF ($Suffix) 
        {
            $NewName = Join-Path $directory $($BaseName+"_"+$TimeStamp+$Extension)
        }
        ELSE
        {
            $NewName = Join-Path $directory $($TimeStamp+"_"+$BaseName+$Extension)
        }
        IF ($CopyItem) {
        Copy-Item -Path $Vars.FullName -Destination $NewName
        }
        ELSE {
        Rename-Item -Path $Vars.FullName -NewName $NewName
        }
        Write-Verbose $("NewName:`t $((Get-ItemProperty $NewName).FullName)")
        return  $((Get-ItemProperty $NewName).FullName)
    }
    End
    {
    }

}
################################################################################
function ConvertTo-PowerShellArray {
    <#
    .SYNOPSIS
    Converts clipboard text or input string into a PowerShell array format.
    
    .DESCRIPTION
    Takes text from the clipboard or a provided string and converts it into a properly formatted
    PowerShell array. Handles both Windows 10 and 11 clipboard formats, and properly processes
    multi-line input including Distinguished Names and other complex strings.
    
    The function now uses a smarter approach to handle clipboard data, ensuring line breaks
    are properly detected regardless of how the operating system provides the clipboard content.
    
    .PARAMETER InputList
    Text to convert into an array. Defaults to clipboard content.
    
    .PARAMETER Sort
    Switch to sort the output in ascending order. Handles both IP addresses and regular text.
    
    .EXAMPLE
    # Multi-line Distinguished Names from clipboard
    # CN=User1,OU=Department,DC=contoso,DC=com
    # CN=User2,OU=Department,DC=contoso,DC=com
    ConvertTo-PowerShellArray
    
    # Comma-separated input
    ConvertTo-PowerShellArray -InputList "item1,item2,item3"
    
    # Space-separated input
    ConvertTo-PowerShellArray -InputList "item1 item2 item3"
    #>
    
    [Alias("CTA", "Format-Array")]
    param (
        [string]$InputList = $(Get-Clipboard -Raw),
        [switch]$Sort
    )

    function Get-BestDelimiter {
        param ([string]$text)
        
        # Test for multiple DN lines
        if (($text -split "\r?\n").Count -gt 1 -and $text -match "^CN=|^OU=|^DC=") {
            return "DN multiline"
        }
        elseif ($text -match "^CN=|^OU=|^DC='") {
            if ($text -match "\s") { return "CN with Spaces" }
            return "CN no spaces"
        }
        elseif (($text -split "\r?\n").Count -gt 2) {
            return "crlf"
        }
        elseif ($text -match ",") {
            return "comma"
        }
        elseif ($text -match "\s") {
            return "space"
        }
        return $null
    }

    $delimiterType = Get-BestDelimiter -text $InputList

    $cleanedList = switch ($delimiterType) {
        "DN multiline" {
            ($InputList -split "\r?\n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
        "crlf" { 
            ($InputList -split "\r?\n") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
        "comma" { 
            ($InputList -split ",") | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
        "space" { 
            ($InputList -split "\s+") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
        default {
            @($InputList)
        }
    }

    if ($Sort) {
        $cleanedList = if ($cleanedList[0] -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
            [string[]]$([version[]]($cleanedList) | Sort-Object -Unique)
        }
        else {
            $cleanedList | Sort-Object -Unique
        }
    }

    $quotedList = $cleanedList | ForEach-Object { "'$($_ -replace "'", "''")'" }
    $result = '@(' + ($quotedList -join ',') + ')'
    
    Write-Output $result
    $result | Set-Clipboard
}

################################################################################
FUNCTION Report-GroupMembers {
	Param([string]$GroupName = $(Read-Host -Prompt "Enter group name"))
	$ErrorActionPreference='SilentlyContinue'
	$collection = Get-ADGroupMember $GroupName | select -ExpandProperty distinguishedName
	foreach ($i in $collection)
    {
        Get-ADUser $i -Properties EmailAddress,whenCreated,accountExpirationDate,LastLogonDate,PasswordExpired,PasswordLastSet,Description -ErrorAction SilentlyContinue | 
            Select-Object -OutVariable +CollectionReport -Property `
                Name,
                SamAccountName,
                EmailAddress,
                whenCreated,
                @{n='AccountExpiration';e={($_.AccountExpirationDate|Get-date).AddDays(-1)}},
                Description,
                Enabled,
                LastLogonDate,
                PasswordExpired,
                PasswordLastSet | 
            out-null
    }
        $CollectionReport | Sort-Object -Property Name | Export-Csv -NoTypeInformation -Path $('.\MembershipReport_'+$GroupName+'_'+$(Get-Date -Format yyy-MM-ddTHHmmss)+'.csv')
}
################################################################################
function Write-Color([String[]]$Text, [ConsoleColor[]]$Color = "White", [int]$StartTab = 0, [int] $LinesBefore = 0,[int] $LinesAfter = 0) {
    $DefaultColor = $Color[0]
    if ($LinesBefore -ne 0) {  for ($i = 0; $i -lt $LinesBefore; $i++) { Write-Host "`n" -NoNewline } } # Add empty line before
    if ($StartTab -ne 0) {  for ($i = 0; $i -lt $StartTab; $i++) { Write-Host "`t" -NoNewLine } }  # Add TABS before text
    if ($Color.Count -ge $Text.Count) {
        for ($i = 0; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine } 
    } else {
        for ($i = 0; $i -lt $Color.Length ; $i++) { Write-Host $Text[$i] -ForegroundColor $Color[$i] -NoNewLine }
        for ($i = $Color.Length; $i -lt $Text.Length; $i++) { Write-Host $Text[$i] -ForegroundColor $DefaultColor -NoNewLine }
    }
    Write-Host
    if ($LinesAfter -ne 0) {  for ($i = 0; $i -lt $LinesAfter; $i++) { Write-Host "`n" } }  # Add empty line after
}

################################################################################

# Generate-RandomPassword generates a random string with rules for placement of specific character classes. 
# In this case it only allows the use of SPECIFIC special characters at char 2-7, and starts/ends with AlphaNum chars

function Generate-RandomPassword {
    param (
        [Parameter(Mandatory=$false)] 
        [string]$Username
    )

    Begin
    {
        function Get-RandomCharacter {
            param ( [char[]]$Characters )
            $index = Get-Random -Minimum 0 -Maximum $Characters.Length
            $Characters[$index]
        }

        # Character classes 
        $charAlpha = @([char[]](65..90) + [char[]](97..122))
        $charNumeric = @([char[]](48..57))
        $charSpecial = @([char[]](33,35,36,37,40,41,43,47,58,61,63))

        $length = Get-Random -Minimum 14 -Maximum 21
    }

    Process
    {
        $password = $null

        while ($password -eq $null) {
            $chars = 1..$length | ForEach-Object {
                if ($_ -eq 1) {
                    # First character (must be a letter)
                    $char = Get-RandomCharacter -Characters $charAlpha
                }
                elseif ($_ -eq $length) {
                    # Last character (must be a letter)
                    $char = Get-RandomCharacter -Characters $charAlpha
                }
                elseif ($_ -ge 2 -and $_ -le 7) {
                    # Special character position
                    $char = Get-RandomCharacter -Characters @($charAlpha + $charSpecial + $charNumeric )
                }
                else {
                    # All other positions (letters or numbers)
                    $char = Get-RandomCharacter -Characters @($charAlpha + $charNumeric)
                }
                $char
            }

            $newPassword = -join $chars

            # Check if the generated password satisfies the criteria
            if ($newPassword -cmatch '[A-Z]' -and $newPassword -cmatch '[a-z]' -and $newPassword -cmatch '\d' -and $newPassword -cmatch '\W' -and $newPassword <# -notlike "*$Username*" #>) {
                $password = $newPassword
            }
        }
    }
    
    End
    {
        $password
    }
}
################################################################################

function Generate-RandomPassPhrase {
<#
.SYNOPSIS
Generates a random passphrase from an external dictionary list of words.

.DESCRIPTION
The Generate-RandomPassPhrase function generates a random passphrase using a word list file. 
The passphrase consists of random words concatenated together with optional padding characters.
The word list file will have one word per line.
By default, the function uses a word list file named "all.words" in the current directory.
The function can also enforce complexity requirements such as case modification and special characters.
Additionally, it supports banning specific words from the generated passphrase.

.PARAMETER minLength
The minimum length of the generated passphrase. Default is 15 characters.

.PARAMETER wordListFile
The path to the word list file. Default is "all.words" in the current directory.

.PARAMETER Iterations
The number of passphrases to generate. Default is 1. Increment to list additional passswords. 

.PARAMETER Complex
Switch parameter to enable passphrase complexity. When specified, modifies case and adds special characters between words.

.PARAMETER BannedWords
An array of words to exclude from the generated passphrase.

.EXAMPLE
Generate-RandomPassPhrase -minLength 20 -wordListFile "custom.words" -Iterations 3 -Complex -BannedWords @("password", "123456")

Generates 3 random passphrases with a minimum length of 20 characters, using a custom word list file "custom.words".
Passphrases are complex with modified case and special characters. The words "password" and "123456" are banned.

.NOTES
This function may require access to external word list files. Ensure appropriate permissions are set.
#>

    param (
        [int]$minLength = 15,
        [string]$wordListFile = 'all.words',
        [int]$Iterations = 1,
        [switch]$Complex, 
        [array]$BannedWords = @('1234','2015','2016','2017','2018','2019','2020','2021','2022','2023','avalanche','broncos','buffalo','buffs','colorado','horse','mammoth','nuggets','rams','rockies','battery','drive','energy','nrel','research','solar','wind','correct','password','qwerty','rain')
    )
        
    #Use a blank space as the character between words, this is modified elsewhere of the Complex switch is used
    $PadCharacters = ' '

    # Initialize an array to store passphrases
    $passphrases = @()

    # Check if the word list file exists
    if (-not (Test-Path -Path $wordListFile -PathType Leaf)) {
        Write-Verbose -Message "Word list file '$wordListFile' not found. Searching for .words files in the current directory..."

        # Search for *.word files in the current directory
        $wordListFiles = Get-ChildItem -Path $pwd -Filter "*.words" -File -Recurse
        Write-Verbose -Message "Found the following words list: $wordListFiles"

        # If no *.words files found, throw an error
        if ($wordListFiles.Count -eq 0) {
            throw "No .words files found in the current directory."
        }

        # Use the first found .word file
        $wordListFile = $wordListFiles[0].FullName
    }

    # Create the wordList as a global variable so that it is not processed everytime the function is run
    if (-not $wordlist) {
        $global:wordList = Get-Content -Path $wordListFile 
    }

#region - PassPhrase generation
    # Generate passphrases for each iteration
    for ($i = 1; $i -le $Iterations; $i++) {
        $password = ''
        $length = 0

        # Keep adding random words until the password length is at least $minLength
        while ($length -lt $minLength) {

            # Get a random word from the list
            $randomWord = Get-Random -InputObject $global:wordList

	        # Check if the $radomWord matches $BannedWord
	        foreach ($bannedWord in $BannedWords) {
		        $isSafe = $true

		        # Use case-insensitive regular expression match
		        if ($randomWord -match [regex]::Escape($bannedWord) -and $matches[0]) {
			        $isSafe = $false
			        break
		        }
	        }

	        # If the word is safe, add it to $password
	        if ($isSafe) {

                # if the Complex switch is used, modify the case
                if ($Complex) {

                    # Special characters to inject, every other char is a space; 50% chance of giving a space. 
                    $PadCharacters = @(' ', ',', ' ', '.', ' ', '/', ' ', '\', ' ', ';', ' ', '-')
                
                    # Generate a random number to determine the case
                    $case = Get-Random -Minimum 0 -Maximum 4

                    # Convert the word to the selected case
                    switch ($case) {
                        0 { $password += $randomWord.ToLower() }                                       # Lower case
                        1 { $password += $randomWord }                                                 # Upper case
                        2 { $password += (Get-Culture).TextInfo.ToTitleCase($randomword.ToLower()) }   # Title case
                        3 { $password += $randomWord.ToUpper() }                                       # No change to case
                    }
                } else {
                    
                    #Add the unmodified word to the password string
                    $password += $randomWord
                }
	        }

            #check the length and continue processing of too short
            $length = $password.Length
            if ($length -lt $minLength) {
            
                # Add a special character, the character set is modified when the Complex switch is used
                $PadChar = Get-Random -InputObject $PadCharacters
                $password += $PadChar

                # Increment length to account for the added space
                $length++
            }
        }

        # Remove trailing space
        $password = $password.TrimEnd()

        # Add the passphrase to the array
        $passphrases += $password
    }
#region - PassPhrase generation

    #Output the password, multiple passwords when the Iterations paramater is used
    $passphrases
}
#####################################################################################
