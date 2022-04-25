## Set Variables:
    $group = "agroup@test.com"
    $members = New-Object System.Collections.ArrayList
    $ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
## Create the Function
    function getMembership($group) {
        $searchGroup = Get-DistributionGroupMember $group -ResultSize Unlimited
        foreach ($member in $searchGroup) {
            if ($member.RecipientTypeDetails-match "Group" -and $member.Alias -ne "") {
                getMembership($member.Alias)
                }           
            else {
                if ($member.Alias -ne "") {
                    if (! $members.Contains($member.Alias) ) {
                        $members.Add($member.Alias) >$null
                        }
                    }
                }
            }
        }

## Run the function
    getMembership($group)

## Output results to screen
    Set-Location $ScriptPath
    $members | Select @{N="UserName";E={$_}} | Export-Csv .\grouplist-rk-d.csv -NTI
