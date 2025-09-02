<#
.SYNOPSIS
    Analyzes group membership patterns to identify potential RBAC roles based on user collections.

.DESCRIPTION
    This script takes a collection of users and analyzes their group memberships to identify
    groups that might represent coherent roles. It calculates inclusivity (how many of your
    users are in each group) and exclusivity (how much of each group consists of your users)
    to help identify meaningful role candidates.
    
    Because apparently someone thought it would be fun to manually copy-paste user queries
    every time they wanted to run this analysis. Well, no more of that nonsense.

.PARAMETER Users
    Array of user identifiers (SamAccountName, DistinguishedName, or ObjectGUID).
    Can be piped in or specified directly.

.PARAMETER SearchBase
    OU distinguished name to search for users. Mutually exclusive with Users parameter.

.PARAMETER Department
    Department filter for users. Can be used with SearchBase.

.PARAMETER Manager
    Manager's distinguished name to filter users. Can be used with SearchBase.

.PARAMETER SecurityGroup
    Security group name to get members from. Mutually exclusive with other parameters.

.PARAMETER MaxGroups
    Maximum number of groups to analyze. Default is 50 because nobody has time for 500+ groups.

.PARAMETER SaveResults
    Automatically save results without prompting. For those who live dangerously.

.EXAMPLE
    .\rbac_analysis_script.ps1 -Users @('jdoe','asmith','bwilson')
    
    Analyzes group membership for the specified users.

.EXAMPLE
    .\rbac_analysis_script.ps1 -SearchBase 'OU=IT,DC=contoso,DC=com' -Department 'Information Technology'
    
    Analyzes all enabled users in the IT OU with IT department.

.EXAMPLE
    'user1','user2','user3' | .\rbac_analysis_script.ps1
    
    Pipeline input because we're fancy like that.

.EXAMPLE
    .\rbac_analysis_script.ps1 -SecurityGroup 'Domain Admins' -MaxGroups 25
    
    Analyzes Domain Admins group members (may cause existential crisis).

.NOTES
    File Name      : rbac_analysis_script.ps1
    Author         : Original concept by some masochist, refactored by Claude (Anthropic)
    Prerequisite   : Active Directory PowerShell module (obviously)
    
    The math here is simple but effective:
    - Inclusivity: What percentage of MY users are in this group?
    - Exclusivity: What percentage of this group consists of MY users?  
    - Final Score: Inclusivity x Exclusivity (sweet spot of relevance)
    
    Groups with high inclusivity but low exclusivity = everyone and their dog has access
    Groups with high exclusivity but low inclusivity = probably too granular
    Groups with high scores for both = potential role goldmine

.LINK
    https://docs.microsoft.com/en-us/powershell/module/activedirectory/
#>

[CmdletBinding(DefaultParameterSetName='UserList')]
param(
    [Parameter(Mandatory=$true, ParameterSetName='UserList', ValueFromPipeline=$true, Position=0)]
    [string[]]$Users,
    
    [Parameter(Mandatory=$true, ParameterSetName='SearchBase')]
    [string]$SearchBase,
    
    [Parameter(ParameterSetName='SearchBase')]
    [string]$Department,
    
    [Parameter(ParameterSetName='SearchBase')]
    [string]$Manager,
    
    [Parameter(Mandatory=$true, ParameterSetName='SecurityGroup')]
    [string]$SecurityGroup,
    
    [Parameter()]
    [int]$MaxGroups = 50,
    
    [Parameter()]
    [switch]$SaveResults
)

begin {
    # Because apparently we need to hold everyone's hand these days
    if (-not (Get-Module -Name ActiveDirectory -ListAvailable)) {
        throw "Active Directory module not found. Install RSAT or run this on a domain controller."
    }
    
    Import-Module ActiveDirectory -ErrorAction Stop
    
    # Collection for pipeline input
    $UserCollection = @()
}

process {
    # Handle pipeline input for users
    if ($PSCmdlet.ParameterSetName -eq 'UserList') {
        $UserCollection += $Users
    }
}

end {
    try {
        Write-Host "`n🔍 Starting RBAC Analysis - Because manual group auditing is for chumps" -ForegroundColor Cyan
        
        #region User Collection Logic
        $UserObjects = @()
        
        switch ($PSCmdlet.ParameterSetName) {
            'UserList' {
                Write-Host "📋 Retrieving user details for $($UserCollection.Count) specified users..." -ForegroundColor Yellow
                foreach ($User in $UserCollection) {
                    try {
                        $UserObj = Get-ADUser -Identity $User -Properties MemberOf,Department,EmployeeType,Manager -ErrorAction Stop
                        $UserObjects += $UserObj
                    }
                    catch {
                        Write-Warning "Failed to retrieve user: $User - $($_.Exception.Message)"
                    }
                }
            }
            
            'SearchBase' {
                Write-Host "🔎 Searching for users in: $SearchBase" -ForegroundColor Yellow
                
                # Build filter based on parameters
                $Filter = "Enabled -eq `$true"
                if ($Department) { $Filter += " -and Department -eq '$Department'" }
                if ($Manager) { $Filter += " -and Manager -eq '$Manager'" }
                
                $UserObjects = Get-ADUser -SearchBase $SearchBase -Filter $Filter -Properties MemberOf,Department,EmployeeType,Manager
            }
            
            'SecurityGroup' {
                Write-Host "👥 Getting members of security group: $SecurityGroup" -ForegroundColor Yellow
                $GroupMembers = Get-ADGroupMember $SecurityGroup -ErrorAction Stop
                foreach ($Member in $GroupMembers) {
                    if ($Member.objectClass -eq 'user') {
                        $UserObj = Get-ADUser -Identity $Member.SamAccountName -Properties MemberOf,Department,EmployeeType,Manager -ErrorAction SilentlyContinue
                        if ($UserObj) { $UserObjects += $UserObj }
                    }
                }
            }
        }
        
        if ($UserObjects.Count -eq 0) {
            throw "No users found matching your criteria. Maybe check your spelling?"
        }
        #endregion User Collection Logic
        
        #region Summary Statistics - The boring but necessary stuff
        $UniqueGroups = $UserObjects.MemberOf | Sort-Object -Unique
        $UniqueDepartments = $UserObjects.Department | Where-Object {$_ -ne $null} | Sort-Object -Unique
        $UniqueManagers = $UserObjects.Manager | Where-Object {$_ -ne $null} | Sort-Object -Unique
        
        Write-Host "`n📊 Analysis Summary:" -ForegroundColor Green
        Write-Host "$($UniqueDepartments.Count)`t`tDepartment Count" -ForegroundColor White
        Write-Host "$($UniqueManagers.Count)`t`tManager Count" -ForegroundColor White  
        Write-Host "$($UserObjects.Count)`t`tUser Count" -ForegroundColor White
        Write-Host "$($UniqueGroups.Count)`t`tTotal Group Memberships" -ForegroundColor White
        
        # Manager breakdown because middle management loves metrics
        if ($UniqueManagers.Count -gt 0) {
            Write-Host "`n👔 Manager's Employee Counts:" -ForegroundColor Green
            foreach ($Manager in $UniqueManagers) {
                $EmployeeCount = ($UserObjects | Where-Object {$_.Manager -eq $Manager}).Count
                $ManagerName = try { (Get-ADUser $Manager).Name } catch { "Unknown Manager" }
                Write-Host "$EmployeeCount`t$ManagerName" -ForegroundColor White
            }
        }
        
        # Department listing
        if ($UniqueDepartments.Count -gt 0) {
            Write-Host "`n🏢 Department List:" -ForegroundColor Green
            $UniqueDepartments | ForEach-Object { Write-Host "  $_" -ForegroundColor White }
        }
        
        # Manager listing  
        if ($UniqueManagers.Count -gt 0) {
            Write-Host "`n👨‍💼 Manager List:" -ForegroundColor Green
            $UniqueManagers | ForEach-Object {
                $ManagerName = try { (Get-ADUser $_).Name } catch { "Unknown: $_" }
                Write-Host "  $ManagerName" -ForegroundColor White
            }
        }
        
        # User listing with employee type because why not be thorough
        Write-Host "`nðŸ‘¤ User List:" -ForegroundColor Green
        $UserObjects | Sort-Object Name | ForEach-Object {
            $EmployeeType = if ($_.EmployeeType) { " ($($_.EmployeeType))" } else { "" }
            Write-Host "  $($_.Name)$EmployeeType" -ForegroundColor White
        }
        #endregion Summary Statistics
        
        #region The Actual Analysis - Where the magic happens
        Write-Host "`n🧮 Crunching numbers for group analysis..." -ForegroundColor Yellow
        
        $UsersInAnalysis = $UserObjects.Count
        $ResultList = @()
        $GroupsToAnalyze = $UniqueGroups | Select-Object -First $MaxGroups
        
        foreach ($Group in $GroupsToAnalyze) {
            try {
                # Users from our collection who are in this group
                $UsersInThisGroup = @($UserObjects | Where-Object {$_.MemberOf -contains $Group})
                
                # Total members of this group (including users not in our analysis)
                $TotalGroupMembers = @(Get-ADGroupMember $Group -ErrorAction Stop).Count
                
                # Skip empty groups because they're useless
                if ($TotalGroupMembers -eq 0) { continue }
                
                # The maths - inclusivity Ã— exclusivity = relevance score
                $InclusivityRatio = $UsersInThisGroup.Count / $UsersInAnalysis
                $ExclusivityRatio = $UsersInThisGroup.Count / $TotalGroupMembers
                $FinalScore = $InclusivityRatio * $ExclusivityRatio
                
                # Get the actual group name instead of the DN because humans
                $GroupName = try { (Get-ADGroup $Group).Name } catch { $Group }
                
                $Result = [PSCustomObject]@{
                    GroupName = $GroupName
                    InCount = $UsersInThisGroup.Count
                    InScore = [Math]::Round($InclusivityRatio, 4)
                    ExCount = $TotalGroupMembers  
                    ExScore = [Math]::Round($ExclusivityRatio, 4)
                    FinalScore = [Math]::Round($FinalScore, 4)
                }
                $ResultList += $Result
            }
            catch {
                Write-Warning "Failed to analyze group: $Group - $($_.Exception.Message)"
            }
        }
        #endregion The Actual Analysis
        
        #region Results Display and Export
        if ($ResultList.Count -eq 0) {
            Write-Warning "No group analysis results generated. This is either very bad or very good."
            return
        }
        
        # Sort by final score descending because the good stuff should be at the top
        $SortedResults = $ResultList | Sort-Object FinalScore -Descending
        
        Write-Host "`n🎯 Top 10 Role Candidates:" -ForegroundColor Green
        $SortedResults | Select-Object -First 10 | Format-Table -AutoSize
        
        # Show the results in GridView for selection
        Write-Host "📋 Opening GridView for detailed analysis and selection..." -ForegroundColor Cyan
        $SelectedResults = $SortedResults | Out-GridView -Title "RBAC Analysis Results - Select groups to export" -OutputMode Multiple
        
        # Handle export logic
        $ShouldSave = $SaveResults -or ($SelectedResults -and (Read-Host "`nSave selected results? (y/N)") -match '^y')
        
        if ($ShouldSave -and $SelectedResults) {
            $Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $BaseName = "RBAC_Analysis_$Timestamp"
            $CsvPath = "$BaseName.csv"
            $TxtPath = "$BaseName.txt"
            
            # Export CSV
            $SelectedResults | Export-Csv -Path $CsvPath -NoTypeInformation
            Write-Host "📊 Results exported to: $CsvPath" -ForegroundColor Green
            
            # Create summary text file because documentation is important
            $SummaryText = @"
RBAC Analysis Summary - $((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
================================================================

Analysis Parameters:
- Parameter Set: $($PSCmdlet.ParameterSetName)
- Total Users Analyzed: $($UserObjects.Count)
- Total Groups Found: $($UniqueGroups.Count)
- Groups Analyzed: $($GroupsToAnalyze.Count)
- Results Selected: $($SelectedResults.Count)

Department Breakdown:
$($UniqueDepartments | ForEach-Object { "  - $_" } | Out-String)

Manager Breakdown:
$(if ($UniqueManagers.Count -gt 0) { 
    $UniqueManagers | ForEach-Object {
        $EmployeeCount = ($UserObjects | Where-Object {$_.Manager -eq $_}).Count
        $ManagerName = try { (Get-ADUser $_).Name } catch { "Unknown: $_" }
        "  - $ManagerName ($EmployeeCount employees)"
    } | Out-String
})

User List:
$(($UserObjects | Sort-Object Name | ForEach-Object {
    $EmployeeType = if ($_.EmployeeType) { " ($($_.EmployeeType))" } else { "" }
    "  - $($_.Name)$EmployeeType"
} | Out-String))

Selected Groups for Export:
$(($SelectedResults | ForEach-Object {
    "  - $($_.GroupName) (Score: $($_.FinalScore), Inclusivity: $($_.InScore), Exclusivity: $($_.ExScore))"
} | Out-String))

Analysis Notes:
- Inclusivity Score: Percentage of analyzed users who are members of this group
- Exclusivity Score: Percentage of group members who are part of the analyzed user set  
- Final Score: Inclusivity Ã— Exclusivity (higher = better role candidate)
- Groups with high final scores represent coherent roles within your user population

Scoring Interpretation:
- 0.8-1.0: Excellent role candidate (most of your users, mostly your users)
- 0.5-0.8: Good role candidate (significant overlap)
- 0.2-0.5: Moderate role candidate (some overlap)
- 0.0-0.2: Weak role candidate (minimal overlap)

Generated by: rbac_analysis_script.ps1 (Refactored by Claude - Anthropic)
"@
            
            $SummaryText | Out-File -FilePath $TxtPath -Encoding UTF8
            Write-Host "🤷 Summary exported to: $TxtPath" -ForegroundColor Green
        }
        elseif (-not $SelectedResults) {
            Write-Host "ðŸ¤· No results selected for export. Your loss." -ForegroundColor Yellow
        }
        #endregion Results Display and Export
        
        Write-Host "`n✅ RBAC Analysis Complete! Now go forth and create some roles." -ForegroundColor Green
    }
    catch {
        Write-Error "Analysis failed spectacularly: $($_.Exception.Message)"
        throw
    }
}