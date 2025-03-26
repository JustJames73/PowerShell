function New-CustomUsername {
    <#
    .SYNOPSIS
    Generates a unique username with the power of regex and AD detective work!
    
    .DESCRIPTION
    Creates a username in the format ex-[FirstInitial][FirstSevenSurnameLetters]
    Handles username collisions like a boss, ensuring uniqueness in Active Directory.
    
    .PARAMETER GivenName
    The first name of the user. We'll steal just the first letter. No mercy!
    
    .PARAMETER SurName
    The last name of the user. We'll chop it down to a svelte 7 letters.
    
    .EXAMPLE
    New-CustomUsername -GivenName "John" -SurName "Doe"
    Returns: ex-jdoe
    
    .EXAMPLE
    New-CustomUsername -GivenName "Alice" -SurName "Wonderland"
    Returns: ex-awonder
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$GivenName,
        
        [Parameter(Mandatory=$true)]
        [string]$SurName
    )

    # Sanitize the input - because computers are picky about their naming conventions
    $firstInitial = $GivenName.Substring(0,1).ToLower()
    $sanitizedSurname = $SurName.Substring(0,([Math]::Min(7, $SurName.Length))).ToLower()

    # Craft the initial username - our username seed prefixed with 'ex-'
    $baseUsername = "ex-$firstInitial$sanitizedSurname"

    # Initiate the username generation quest!
    $finalUsername = $baseUsername
    $counter = 1

    # The Great Username Collision Detector™
    while (Get-ADUser -Filter {SamAccountName -eq $finalUsername} -ErrorAction SilentlyContinue) {
        # If we've reached the max length of 11 characters, time for some username surgery
        if ($finalUsername.Length -eq 11) {
            $finalUsername = $baseUsername.Substring(0,10) + $counter.ToString()
        }
        else {
            $finalUsername = $baseUsername + $counter
        }
        
        # Increment our counter - the number of failed username attempts
        $counter++
    }

    return $finalUsername
}

# Example usage (commented out - uncomment to test)
# $newUsername = New-CustomUsername -GivenName "John" -SurName "Doe"
# Write-Host "Generated Username: $newUsername"
