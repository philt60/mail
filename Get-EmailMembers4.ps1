# Phils script to expand emails to their constituents (user, group or contact) from WIN  AD.
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
    Begin {
        Write-Verbose "emails: $emails"
        $searcher = New-Object DirectoryServices.DirectorySearcher
        $searcher.SearchRoot = [adsi]'LDAP://OU=Moira,DC=WIN,DC=MIT,DC=EDU'
    }
    Process {
        foreach ($email in $emails) {
            Write-Verbose "email: $email"
            $searcher.Filter = "(mail=$email)"
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
                Write-Warning("Search error: $_")
                continue
            }

            foreach ($res in $results) {
                $ADEntry = $res.GetDirectoryEntry()
                $members = [System.Collections.Generic.List[System.String]]::new()
                if ($ADEntry.Properties.objectclass -contains "group") {  # Check if the result is a group
                    Write-Verbose("`tMembers of the group: ")
                    if (! $ADEntry.Properties.Contains("member")) {
                        Write-Host("{0} has no members" -f $($ADEntry.Properties.displayname)) -ForegroundColor Magenta
                        [PSCustomObject]@{
                            Email   = $email
                            Class   = "group"
                            Members = $null
                        }
                        continue
                    }
                    foreach ($member in $ADEntry.Properties.member) {
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
                } elseif ($ADEntry.Properties.objectclass -contains "user") {
                    Write-Verbose("`tMembers of User: ")
                    $memberValue = switch ($ReturnType) {
                        "DisplayName" { $ADEntry.Properties.displayname }
                        "Email" { $ADEntry.Properties.mail }
                        "AccountName" { $ADEntry.Properties.sAMAccountName }
                    }
                    $class = "user"
                    $members.Add($memberValue)
                } elseif ($ADEntry.Properties.objectclass -contains "contact") {
                    Write-Verbose("`tMembers of contact: ")
                    $memberValue = switch ($ReturnType) {
                        "DisplayName" { $ADEntry.Properties.displayname }
                        "Email" { $ADEntry.Properties.mail }
                        "AccountName" { $ADEntry.Properties.sAMAccountName }
                    }
                    $class = "contact"
                    $members.Add($memberValue)
                } else {
                    Write-Warning "Alien ObjectClass has been found: $($ADEntry.Properties.objectclass)"
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
    end {
        $searcher.Dispose()
    }
}
# "phils@mit.edu", "prt_adm@mit.edu", "sts.ios@mit.edu" | Get-EmailMembers -verbose
# Get-EmailMembers -emails "phils@mit.edu", "prt_adm@mit.edu", "sts.ios@mit.edu", "phils" -verbose -ReturnType DisplayName


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

### eof ###