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
