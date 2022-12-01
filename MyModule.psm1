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
the whole line of text matching the search
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
<#
Consider testing the input for common delimiters and using the split method
Add an option to sort the output
Add an option to send the output to clipboard
Return the processed value as standard output
#>

function Format-Array 
{
<#
.SYNOPSIS
Paste a line separated list at the prompt to be re-formatted into the @('val1','val2') array format.
.DESCRIPTION
The Format-Array function will take a line separated list, such as one copied 
from a spreadsheet, and rearrange it into a quoted & comma separated list 
compatible with the PowerShell array. The array is copied to the clipboard. 
.EXAMPLE
	C:\> Format-Array
	Paste list
	: Val1
	Val2
	Val3
.PARAMETER <>
None
 #>
    $C0 = ($(Read-Host -Prompt "paste list`n") -isplit "`n");
    foreach ($I in $C0) { $C1 += ([string]::Concat("`'"+$(($i).TrimStart().TrimEnd())+"`',")) };
    $C2 = @($C1.Substring(0,$($C1.Length-1))); 
    $C2 = ([string]::Concat("@`(" + $C2 + "`)")); 
    $C2 | clip.exe
    
}
################################################################################
FUNCTION Report-GroupMembers {
	Param([string]$GroupName = $(Read-Host -Prompt "Enter group name"))
	$ErrorActionPreference='SilentlyContinue'
	$collection = Get-ADGroupMember $GroupName | select -ExpandProperty distinguishedName
	foreach ($i in $collection)
    {
        Get-ADUser $i -Properties EmailAddress,whenCreated,accountExpirationDate,LastLogonDate,PasswordExpired,PasswordLastSet -ErrorAction SilentlyContinue | Select-Object -Property Name,SamAccountName,EmailAddress,whenCreated,@{n='AccountExpiration';e={($_.AccountExpirationDate|Get-date).AddDays(-1)}},LastLogonDate,PasswordExpired,PasswordLastSet -OutVariable +CollectionReport | out-null
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

function Get-Excuse
{
  $url = 'http://pages.cs.wisc.edu/~ballard/bofh/bofhserver.pl'
  $ProgressPreference = 'SilentlyContinue'
  $page = Invoke-WebRequest -Uri $url -UseBasicParsing
  $pattern = '(?m)<br><font size = "\+2">(.+)'
  if ($page.Content -match $pattern)
  {
    $matches[1]
  }
}
