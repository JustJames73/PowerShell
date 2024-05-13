function ConvertTo-PowerShellArray {
<#
.SYNOPSIS
Use text in the clipboard to create a powershell array value in plain text. 
.DESCRIPTION
The -IndputList can be a space, comma or line separated list of values that can 
be copied from a spreasheet and rearrange it into a quoted & comma separated list. 
The -Sort switch can be used to properly sort text or IP addresses in ascending order. 
The output is sent both to the screen and the clipboard for pasting. 
.ALIAS
    CTA
    Format-Array
.EXAMPLE
    $inputList = 'thing1 thing2 thing3'
    $result = ConvertTo-PowerShellArray -InputList $inputList
    $result  # Outputs the formatted array

    # Alternatively, you can use the aliases "CTA" or "Format-Array"
    $result2 = CTA -InputList $inputList
    $result2  # Outputs the formatted array

    $result3 = Format-Array -InputList $inputList
    $result3  # Outputs the formatted array

    $inputList = '192.168.10.10 192.168.2.1 192.168.1.100'
    $result4 = ConvertTo-PowerShellArray -InputList $inputList -Sort
    $result4  # Outputs the formatted array with sorted IP addresses
.PARAMETER <>
    -InputList (Parameter)
        Use clipboard content by default or specify as a quoted string. 
    -Sort (Switch)
        Sort in ascending order, IPv4 addresses are detected and 
        sorted appropriately. 
 #>


    #[Alias("CTA", "Format-Array")]
    param (
        [string]$InputList = $(Get-Clipboard) ,
        [switch]$Sort
    )

   # Determine the delimiter based on the format of the input list
   if ($InputList -match ',') {
        $delimiter = ','
    }
    elseif ($InputList -match '\r?\n') {
        $delimiter = '\r?\n'
    }
    else {
        $delimiter = ' '
    }

    # Trim blank spaces, remove blank lines
    $cleanedList = ($InputList -split $delimiter).Trim() | Where-Object { $_.Trim() -ne '' }

    # Sort the IP address list in ascending order if the -Sort switch is used
    if ($Sort) {
        $cleanedList = $(
            if ($cleanedList[0] -match '\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}') { [string[]]$([version[]]($cleanedList) | Sort-Object -Unique) }
            else { $cleanedList | Sort-Object -Unique }
            )
    }

    # Add quotes to each string
    $cleanedList = $cleanedList | ForEach-Object { "'$_'" }

    # Join the cleaned list with commas and format as PowerShell array
    $result = '@(' + ($cleanedList -join ',') + ')'

    # Output result to console
    Write-Output $result

    # Copy result to clipboard
    $result | Set-Clipboard
}

# Example usage:
# ConvertTo-PowerShellArray -InputList "value1, value2, value3"
