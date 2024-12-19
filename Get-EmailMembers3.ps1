# Phils script to expand an emails to their constituent user, group or contact in Win  AD.
#
Set-StrictMode -Version Latest
Set-Location $PSScriptRoot

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
            $searcher.Filter = "(mail=$email)"
            $searcher.SearchRoot = [adsi]'LDAP://OU=Moira,DC=WIN,DC=MIT,DC=EDU'
            Write-Verbose("Searching {0} in {1}" -f $searcher.Filter, $searcher.SearchRoot.Path)
            try {
                if (!($results = @($Searcher.FindAll()))) {
                    Write-Warning("No results found for $email")
                    [PSCustomObject]@{
                        Email   = $email
                        Class   = ""
                        Members = ("NOT FOUND IN AD")
                    }
                    continue
                }
            } catch {
                Write-Warning("$_")
                continue
            }
            $searcher.Dispose()

            foreach ($res in $results) {
                $entry = $res.GetDirectoryEntry()
                # Write-Host("res entry {0}" -f $($entry.Properties.displayname)) -ForegroundColor Yellow
                $members = [System.Collections.Generic.List[System.String]]::new()
                if ($entry.Properties.objectclass -contains "group") {  # Check if the result is a group
                    Write-Verbose("`tMembers of the group: ")
                    if (! $entry.Properties.Contains("member")) {
                        Write-Host("{0} has no members" -f $($entry.Properties.displayname)) -ForegroundColor Magenta
                        [PSCustomObject]@{
                            Email   = $email
                            Class   = "group"
                            Members = $null
                        }
                        continue
                    }
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
                        $class = "group"
                        $members.Add($memberValue)
                    }
                } elseif ($entry.Properties.objectclass -contains "user") {
                    Write-Verbose("`tMembers of User: ")
                    $memberValue = switch ($ReturnType) {
                        "DisplayName" { $entry.Properties.displayname }
                        "Email" { $entry.Properties.mail }
                        "AccountName" { $entry.Properties.sAMAccountName }
                    }
                    $class = "user"
                    $members.Add($memberValue)
                } elseif ($entry.Properties.objectclass -contains "contact") {
                    Write-Verbose("`tMembers of contact: ")
                    $memberValue = switch ($ReturnType) {
                        "DisplayName" { $entry.Properties.displayname }
                        "Email" { $entry.Properties.mail }
                        "AccountName" { $entry.Properties.sAMAccountName }
                    }
                    $class = "contact"
                    $members.Add($memberValue)
                } else {
                    Write-Warning "Alien ObjectClass has been found: $($entry.Properties.objectclass)"
                    continue
                }
                Write-Verbose "`t$members"
                [PSCustomObject]@{
                    Email   = $email
                    Class   = $class
                    Members = ($members -join ", ")
                }
            }
        }
    }
}

# Import the CSV file
$aliases = Import-Csv -Path "aliases.csv"
# Delete the Excel file before writing to it
$outputExcel = "Members.xlsx"
Remove-Item $outputExcel -Force -ErrorAction SilentlyContinue

# Process each row in the CSV
foreach ($row in $aliases) {
    $emails = $row.Emails -split ','
    $processedEmails = @()
    foreach ($email in $emails) {
        $email = $email.Trim()
        if ($email -in @("root", "/dev/null")) {
            continue
        }
        if ($email -inotmatch '^[\w\.\+\-]+@[\w\.-]+\.\w+$') {     # not an email, search in name column to see if forwarded.
            if ($forwarded = ($aliases |? Name -eq $email)) {
                if ($forwarded.emails -match '^[\w\.\+\-]+@mit.edu$' ) { 
                    $email = $forwarded.Emails
                } else {
                    Write-Host "Forwarding $email > $forwarded.Emails" -ForegroundColor Red
                }
            } else {
                $email = $email + "@sloan.mit.edu"   # best guess for address
           }
        } 
        $processedEmails += $email
    }
    if (! $processedEmails) {
        continue
    }
    # Call the Get-EmailMembers function with the processed emails
    Write-Verbose "$($row.Name) = $processedEmails"
    $results = Get-EmailMembers -emails $processedEmails -ReturnType Email -Verbose
    # Add the "name" value to the results and make it the first column
    $finalResults = foreach ($result in $results) {
        [PSCustomObject]@{
            Name    = $row.Name
            Email   = $result.Email
            Class   = $result.Class
            Members = $result.Members
        }
    }
    # Export the results to Excel
    if (-not (Test-Path $outputExcel)) {
        $finalResults | Export-Excel -Path $outputExcel -WorksheetName "Members"
    } else {
        $finalResults | Export-Excel -Path $outputExcel -WorksheetName "Members" -Append
    }
}