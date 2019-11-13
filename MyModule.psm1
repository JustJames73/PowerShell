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
                  HelpUri = 'http://www.microsoft.com/',
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
FUNCTION GetADUserByEnum ($employeeNumber,[switch]$RefreshUsers) {
	BEGIN {
		IF ($ArrayOfUserByEnum) {
            IF ($RefreshUsers) {
                $global:ArrayOfUserByEnum = Get-ADUser -Properties AccountExpirationDate,Department,Description,mail,EmployeeNumber,GivenName,Surname,Manager,SamAccountName,Office,DisplayName,cwNumber,EmployeeID,employeeType,accountType,expDate,LeftNREL,whencreated,NRELstartDate -Filter 'employeeNumber -like "*"' 
            }
        } ELSE { 
			$global:ArrayOfUserByEnum = Get-ADUser -Properties AccountExpirationDate,Department,Description,mail,EmployeeNumber,GivenName,Surname,Manager,SamAccountName,Office,DisplayName,cwNumber,EmployeeID,employeeType,accountType,expDate,LeftNREL,whencreated,NRELstartDate -Filter 'employeeNumber -like "*"' 
		}
	}
	PROCESS {
		IF ($employeeNumber -in $ArrayOfUserByEnum.EmployeeNumber) {
	        RETURN $($ArrayOfUserByEnum | Where-Object employeenumber -like $employeeNumber)
	    } 
	}
	END {}
}
################################################################################
function Report-GroupMembers {
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
FUNCTION New-Profile {
New-Item -ItemType File -Path $profile -Force
notepad $profile
}
################################################################################
FUNCTION Start-ExchangeShell {

    IF ((Get-PSSession).ConfigurationName -notmatch "Microsoft.Exchange")
        # Check to see if the Exchange session has already been started. 
    { 
        $ExchangeFQDN = 'http://xp11mbx3.nrel.gov/powershell/'; 
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeFQDN -Authentication Kerberos; 
        Import-PSSession $Session -AllowClobber -WarningAction SilentlyContinue | Out-Null
        # -CommandName Get-MailboxExportRequest,New-MailboxExportRequest,Remove-MailboxExportRequest,Get-MailContact,Disable-MailContact,Get-MailUser,Disable-MailUser,Get-MailboxStatistics 
    }
    ELSE
    {
    "The Microsoft Exchange shell commands are already loaded."
    }
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

################################################################################
<# Functions to work on ... 
FUNCTION Report-ServiceNow 
{
    Param([string[]]$SAM = $(Read-Host -Prompt 'Enter user name'))
    BEGIN
    {
        $ErrorView = "CategoryView"; 
        $Accounts = foreach ($S in $SAM) {@{SamAccountName=$($S)}}
    }
    PROCESS
    {
        $Accounts | % {
            $UserProp = $(Get-ADUser $_.SamAccountName -Properties MemberOf,Mail,WhenCreated,AccountExpirationDate,Department,info );`
            if ($UserProp.info -ne $null) { 
                $RITM = $($UserProp.info)
            } else {
                $RITM = "No RITM"
            }
            if ($UserProp.Department -ne $null) {
                $Department = $($UserProp.Department)
            } else {
                $Department = "N/A"
            }
            if ($UserProp.AccountExpirationDate -ne $null) {
                $AccountExpires = $(($UserProp.AccountExpirationDate | Get-Date).AddDays(-1) | Get-Date -Format D)
            } else {
                $AccountExpires = "N/A"
            }
         ;`
        $NTGroups = $($UserProp.MemberOf | Get-ADGroup | Sort-Object -Property Name | select -ExpandProperty Name );`
        $eMailAddress = $UserProp.Mail;`
        $AccountCreated = $($UserProp.WhenCreated | Get-Date -Format D);`
        $UserName = $UserProp.SamAccountName ;`
        $DisplayName = $UserProp.Name ;`
		$NEWUID = $((Get-ADUser $_.SamAccountName -Properties uidNumber).uidnumber)
write-host $(
"$RITM `
Created account for [ $DisplayName ]`
Account Created: [ $AccountCreated ]`
Account End Date: [ $AccountExpires ]`
Center Number: [ $Department ]`
NT Groups: [ $(($NTGroups) -join ", ") ]`
E-Mail Address: [ $eMailAddress ]`
User Name: [ $UserName ] `
Unix UidNumber: [$($NEWUID)] `
Please contact the Service Operations Center at 303-275-4171 or Service.Center@nrel.gov for log-on assistance.`
==============================================================================================================`n"
);`
      }
    }
    END
    {}
}

################################################################################
Try { 
  if ((Get-PSSnapin | where { $_.Name -eq "Quest.ActiveRoles.ADManagement" }) -eq $null) 
  { Add-PSSnapin Quest.ActiveRoles.ADManagement } 
}
Catch [System.Exception] { 
  Write-Host "Quest ActiveRoles ADManagement is already loaded" 
}
Try { 
  if ((Get-Module | where { $_.Name -eq "ActiveDirectory" }) -eq $null) 
  { Import-Module ActiveDirectory } 
}
Catch [System.Exception] { 
  Write-Host "ActiveDirectory Module is already loaded" 
}

################################################################################
# FUNCTIONS ##################### FUNCTIONS ######################## FUNCTIONS #
################################################################################
$FunctionsRef = ls function: | select Name
################################################################################
$ExistingUsers = Get-ADUser -Properties employeenumber,displayname,mail -Filter 'employeeNumber -like "*"' 
FUNCTION GetADUserByEnum ($employeeNumber,$SearchBase) {
    IF ($employeeNumber -in $ExistingUsers.EmployeeNumber) {
        RETURN $($ExistingUsers | Where-Object employeenumber -like $employeeNumber)
    } 
}
################################################################################
FUNCTION Search-Scripts {
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
################################################################################
FUNCTION Prompt { 
  (Get-Host).ui.rawui.foregroundcolor = "Green"
  "PS [$env:COMPUTERNAME] [$((Get-ItemProperty $pwd).BaseName)] >"
}
################################################################################
FUNCTION New-Password {
  param ($name)
  trap [Microsoft.ActiveDirectory.Management.Commands.GetADUser] {"ADIdentityNotFoundException"; break;}
  if ($name -eq $null) { 
    Write-Host $(Get-Date -Format "MMM\@dd\#yy\$")
  }
  else {     
    $P = (Get-ADUser $name -Properties whenCreated | Select-Object -ExpandProperty whenCreated | Get-Date -Format "MMM\@dd\#yy\$")
    $P
  }
}
#Get-ADUser $name | Set-ADAccountPassword -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $P -Force)
################################################################################
FUNCTION Add-TimeStamp {
  Param(
    [Parameter(
        Mandatory=$True,
        HelpMessage="Enter the filename to prefix with a sortable timestamp. Use the Switch -Suffix to add the timestamp to the end of the filename"
    )]
    [string] $Path,
    [switch] $Suffix,
    [switch] $CopyItem
  )
  Write-Verbose -Message $("OldName:`t$((Get-ItemProperty $Path).FullName)")
  $Vars = Get-ItemProperty $Path | select BaseName,Directory,Extension,LastWriteTime
  $Directory = $([string]($vars | select -ExpandProperty Directory)+'\')
  $BaseName = ($vars | select -ExpandProperty BaseName)
  $Extension = ($vars | select -ExpandProperty Extension)
  $Time = ($vars | select -ExpandProperty LastWriteTime)
  $TimeStamp = ($Time | Get-Date -Format yyy-MM-ddTHHmmss)
  IF ($Suffix) 
  {
    $NewName = ($BaseName+"_"+$TimeStamp+$Extension)
  }
  else
  {
    $NewName = ($TimeStamp+"_"+$BaseName+$Extension)
  }
  $Path1 = $([string]($Directory+$BaseName+$Extension))
  IF ($CopyItem) {
    Copy-Item -Path $Path1 -Destination $NewName
  }
  ELSE {
    Rename-Item -Path $Path1 -NewName $NewName
  }
  Write-Verbose $("NewName:`t $((Get-ItemProperty $NewName).FullName)")
  return  $((Get-ItemProperty $NewName).FullName)
}

################################################################################
FUNCTION MakeArray 
{
    $Collection = ($(Read-Host -Prompt 'paste list') -isplit "`n");
    foreach ($I in $Collection) { $Collection2 += ([string]::Concat("`""+$i+"`",")) };
    $Collection1 = @($Collection2.Substring(0,$($Collection2.Length-1))); 
    $Collection1 = ([string]::Concat("`$Collection`=`@`(" + $Collection1 + "`)")); 
    Write-Host $Collection1 
}
################################################################################
function Report-ServiceNow { 
    Param([string[]]$SAM = $(Read-Host -Prompt 'Enter user name'))
    BEGIN
    {
        $ErrorView = "CategoryView"; 
        $Accounts = foreach ($S in $SAM) {@{SamAccountName=$($S)}}
    }
    PROCESS
    {
        $Accounts | % {
            $UserProp = $(Get-ADUser $_.SamAccountName -Properties MemberOf,Mail,WhenCreated,AccountExpirationDate,Department,info,uidnumber );`
            if ($UserProp.info -ne $null) { 
                $RITM = $($UserProp.info)
            } else {
                $RITM = "No RITM"
            }
            if ($UserProp.Department -ne $null) {
                $Department = $($UserProp.Department)
            } else {
                $Department = "N/A"
            }
            if ($UserProp.AccountExpirationDate -ne $null) {
                $AccountExpires = $(($UserProp.AccountExpirationDate | Get-Date).AddDays(-1) | Get-Date -Format D)
            } else {
                $AccountExpires = "N/A"
            }
         ;`
        $NTGroups = $($UserProp.MemberOf | Get-ADGroup | Sort-Object -Property Name | select -ExpandProperty Name );`
        $eMailAddress = $UserProp.Mail;`
        $newUID = $UserProp.uidnumber; `
        $AccountCreated = $($UserProp.WhenCreated | Get-Date -Format D);`
        $UserName = $UserProp.SamAccountName ;`
        $DisplayName = $UserProp.Name ;`
write-host $(
"$RITM `
Created account for [ $DisplayName ]`
Account Created: [ $AccountCreated ]`
Account End Date: [ $AccountExpires ]`
Center Number: [ $Department ]`
NT Groups: [ $(($NTGroups) -join ", ") ]`
E-Mail Address: [ $eMailAddress ]`
User Name: [ $UserName ] `
Unix UidNumber: [$newUID] `
Please contact the Service Operations Center at 303-275-4171 or Service.Center@nrel.gov for log-on assistance.`
==============================================================================================================`n"
);`
      }
    }
    END
    {}
}  
################################################################################
FUNCTION Restart-ADAudit 
{
	get-service -ComputerName XADAUDIT -Name 'ADAudit Plus' `
	| Set-Service -Status Running
}
################################################################################

################################################################################
cd $env:USERPROFILE; Write-Host $(pwd); powershell_ise.exe
################################################################################
# $FunctionsRef = ls function: | select Name ## This variable is defined at the begining of the profile before any other functions are defined. 
$FunctionsDif = ls function: | select Name
$FunctionsRes = compare $FunctionsRef $FunctionsDif -Property Name | ? SideIndicator -like '=>' | select -ExpandProperty Name
Write-Host -BackgroundColor Black -ForegroundColor Cyan "Your custom functions are:`n $FunctionsRes"
################################################################################

################################################################################
<#
# Retired functions
# FUNCTION Get-MailboxCount { Get-mailboxserver | Get-MailboxDatabase | Where-Object -FilterScript { $($_.Identity -like "MAILBOX*" -and $_.Identity -notlike '*mbx1sg1' -and $_.Identity -notlike  '*interns*' -and $_.Identity -notlike '*Resource*') } | select Identity,@{n='MailboxCount';e={@(Get-Mailbox -database $_.identity).count}} | sort MailboxCount | Out-GridView }
################################################################################
function DeDupe ($param1)
{
    [string[]]$param1
    Get-Content $param1 | Sort-Object | Get-Unique
}
################################################################################
function Add-TimeStamp {
  Param(
    [Parameter(
        Mandatory=$True,
        HelpMessage="Enter the filename to prefix with a sortable timestamp. Use the Switch -Suffix to add the timestamp to the end of the filename"
    )]
    [string] $Path,
    [switch] $Suffix
  )
  Write-Host "OldName:`t" $((Get-ItemProperty $Path).FullName)
  $Vars = Get-ItemProperty $Path | select BaseName,Directory,Extension,LastWriteTime
  $Directory = $([string]($vars | select -ExpandProperty Directory)+'\')
  $BaseName = ($vars | select -ExpandProperty BaseName)
  $Extension = ($vars | select -ExpandProperty Extension)
  $Time = ($vars | select -ExpandProperty LastWriteTime)
  $TimeStamp = ($Time | Get-Date -Format yyy-MM-ddTHHmmss)
  if ($Suffix) 
  {
    $NewName = ($BaseName+"_"+$TimeStamp+$Extension)
  }
  else
  {
    $NewName = ($TimeStamp+"_"+$BaseName+$Extension)
  }
  $Path1 = $([string]($Directory+$BaseName+$Extension))
  Rename-Item -Path $Path1 -NewName $NewName
  Write-Host "NewName:`t" $((Get-ItemProperty $NewName).FullName)
}
################################################################################
#>
