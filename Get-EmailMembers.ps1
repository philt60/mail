Set-StrictMode -Version Latest

Function Get-EmailMembers {
    [CmdletBinding()]
    [OutputType([System.Collections.Generic.List[System.String]])]
    Param (
        [Parameter(Mandatory, Position=0, ValueFromPipeline=$true, HelpMessage="Enter SMTP addresses to search for in Active-Directory.")][string[]]$emails,
        [Parameter(HelpMessage="Specify the type of information to return: DisplayName, Email, or AccountName.")][ValidateSet("DisplayName", "Email", "AccountName")][string]$ReturnType = "AccountName"
    )

    Process {
        $membersList = [System.Collections.Generic.List[System.String]]::new()
        foreach ($email in $emails) {
            Write-Verbose "email: $email"
            $searcher = New-Object DirectoryServices.DirectorySearcher
            $searcher.Filter = "(mail=$email)"
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

                # Check if the result is a group
                if ($entry.Properties.objectclass -contains "group") {
                    Write-Host("Members of the group:") -ForegroundColor Yellow
                    foreach ($member in $entry.Properties.member) {
                        $memberEntry = [adsi]"LDAP://$member"
                        switch ($ReturnType) {
                            "DisplayName" { $membersList.Add($memberEntry.displayname) }
                            "Email" { 
                                if ($memberEntry.mail) {
                                    $membersList.Add($memberEntry.mail)
                                } else {
                                    $membersList.Add($memberEntry.cn)
                                }
                            }
                            "AccountName" { 
                                if ($memberEntry.sAMAccountName) {
                                    $membersList.Add($memberEntry.sAMAccountName)
                                } else {
                                    $membersList.Add($memberEntry.cn)
                                }
                            }
                        }
                    }
                } else {
                    # Return the specified value for users
                    switch ($ReturnType) {
                        "DisplayName" { $membersList.Add($entry.Properties.displayname) }
                        "Email" { 
                            if ($entry.Properties.mail) {
                                $membersList.Add($entry.Properties.mail)
                            } else {
                                $membersList.Add($entry.Properties.cn)
                            }
                        }
                        "AccountName" { 
                            if ($entry.Properties.sAMAccountName) {
                                $membersList.Add($entry.Properties.sAMAccountName)
                            } else {
                                $membersList.Add($entry.Properties.cn)
                            }
                        }
                    }
                }
            }
        }
        return $membersList
    }
}

Get-EmailMembers "emba12.mitsloan@mit.edu", "phils@mit.edu" -ReturnType DisplayName -Verbose

# Get-EmailMembers "emba12.mitsloan@mit.edu" -ReturnType AccountName
#  "sts.staff.mitsloan@mit.edu", "emba12.mitsloan@mit.edu", "phils@mit.edu" | Get-EmailMembers -Verbose