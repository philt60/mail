Set-StrictMode -Version Latest

Function Get-EmailMembers {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param (
        [Parameter(Mandatory, Position=0, ValueFromPipeline=$true, HelpMessage="Enter SMTP addresses to search for in Active-Directory.")][string[]]$emails,
        [Parameter(HelpMessage="Specify the type of information to return: DisplayName, Email, or AccountName.")][ValidateSet("DisplayName", "Email", "AccountName")][string]$ReturnType = "AccountName"
    )

    Process {
        foreach ($email in $emails) {
            Write-Verbose "email: $email"
            $searcher = New-Object DirectoryServices.DirectorySearcher
            $searcher.Filter = "(anr=$email)"
            $searcher.SearchRoot = [adsi]'LDAP://OU=Moira,DC=WIN,DC=MIT,DC=EDU'
            Write-Verbose("Searching {0} in {1}" -f $searcher.Filter, $searcher.SearchRoot.Path)
            try {
                if (!($results = @($Searcher.FindAll()))) {
                    Write-Warning("No results found for $email")
                    continue
                }
            } catch {
                Write-Warning("$_")
                continue
            }
            $searcher.Dispose()

            foreach ($res in $results) {
                $entry = $res.GetDirectoryEntry()
                Write-Host("{0}" -f $($entry.Properties.displayname)) -ForegroundColor Cyan

                $members = [System.Collections.Generic.List[System.String]]::new()
                # Check if the result is a group
                if ($entry.Properties.objectclass -contains "group") {
                    Write-Host("Members of the group:") -ForegroundColor Yellow
                    foreach ($member in $entry.Properties.member) {
                        $memberEntry = [adsi]"LDAP://$member"
                        $memberValue = switch ($ReturnType) {
                            "DisplayName" { $memberEntry.displayname }
                            "Email" { 
                                if ($memberEntry.mail) {
                                    $memberEntry.mail
                                } else {
                                    $memberEntry.cn
                                }
                            }
                            "AccountName" { 
                                if ($memberEntry.sAMAccountName) {
                                    $memberEntry.sAMAccountName
                                } else {
                                    $memberEntry.cn
                                }
                            }
                        }
                        $members.Add($memberValue)
                    }
                } else {
                    # Return the specified value for users
                    $memberValue = switch ($ReturnType) {
                        "DisplayName" { $entry.Properties.displayname }
                        "Email" { $entry.Properties.mail }
                        "AccountName" { $entry.Properties.sAMAccountName }
                    }
                    $members.Add($memberValue)
                }
                [PSCustomObject]@{
                    Email  = $email
                    Members = ($members -join ", ")
                }
            }
        }
    }
}

<# 
$outputExcel = "Members.xlsx"
Remove-Item  $outputExcel -Force -ea SilentlyContinue
# Get-EmailMembers "emba12.mitsloan@mit.edu", "phils@mit.edu" -ReturnType DisplayName -Verbose | Export-Excel -Path "Members.xlsx" -WorksheetName "Members"
# Get-EmailMembers "emba12.mitsloan@mit.edu" -ReturnType AccountName
# "sts.staff.mitsloan@mit.edu", "emba12.mitsloan@mit.edu", "phils@mit.edu", "stacyp", "lisa.farrell" | Get-EmailMembers | Export-Excel -Path "Members.xlsx" -WorksheetName "Members" -Show
"emba12.mitsloan@mit.edu", "phils@mit.edu", "stacyp", "lpradell" | Get-EmailMembers -ReturnType Email | Export-Excel -Path $outputExcel -WorksheetName "Members" -Show
#>


# Import the CSV file
$aliases = Import-Csv -Path "aliases.csv"

# Function to check if a string is an email address
Function IsEmail {
    param (
        [string]$str
    )
    return $str -match '^[\w\.-]+@[\w\.-]+\.\w+$'
}

# Process each row in the CSV
foreach ($row in $aliases) {
    $emails = $row.Emails -split ','
    $processedEmails = @()

    foreach ($email in $emails) {
        $email = $email.Trim()
        if (-not (IsEmail -str $email)) {
            $email = "$email@mit.edu"
        }
        $processedEmails += $email
    }

    # Call the Get-EmailMembers function with the processed emails
    Get-EmailMembers -emails $processedEmails -ReturnType Email # | Export-Excel -Path "Members.xlsx" -WorksheetName "Members" -Show
}
