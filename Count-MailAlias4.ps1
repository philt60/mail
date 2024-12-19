#!/usr/bin/env pwsh
# This reads in maillogs and counts successfully forwarded mail and usage by alias lists.
# -phils 04/2023

param(
    [Parameter(Mandatory=$False, Position=0)][DateTime]$startDate = ((Get-Date).Date.AddDays(-90)),
    [Parameter(Mandatory=$False, Position=1)][DateTime]$endDate = (Get-Date).Date
)

Set-StrictMode -Version Latest

# Define directories based on OS
if ($env:OS -eq "Windows_NT") {
    $aliasDir = "c:\tmp\mail7"     # windows practice directory, usually /etc/postfix
    $logDir = "C:\tmp\mail7"       # windows practice directory, usually /var/log/
    $outputCsv = "c:\tmp\mail7\mail_alias_counts.csv"
} else {
    $aliasDir = "/etc/postfix"        # where the aliases are, usually /etc/postfix
    $logDir = "/var/log/maillogs"       # where the maillogs are, usually /var/log/
    $outputCsv = "/var/tmp/mail_alias_counts.csv"
}

# Construct alias paths
$aliasPaths = @("aliases", "alias_clb_off") | % { Join-Path -Path $aliasDir -ChildPath $_ -Resolve -ErrorAction Stop }
$aliases = Get-Content -Path $aliasPaths -ea Stop | % { if ($_ -match '^(?!#)(?<name>\S+)\s*:\s*(?<list>\S+.*)$') { $Matches.name}}
Write-Verbose("$aliasDir aliases: {0}" -f $aliases.Count) -Verbose

$totalLines = 0
Remove-Item -LiteralPath $outputCsv -Force -ErrorAction SilentlyContinue

# Process log files
$myCsv = foreach ($logFile in Get-ChildItem -LiteralPath $logDir -Filter "maillog-20*" | ? {$_.LastWriteTime -gt $startDate -and $_.LastWriteTime -lt $endDate } ) {
    Write-Verbose ("{0}" -f $logFile) -Verbose
    if (!($date = [datetime]::ParseExact($($logFile.BaseName -Replace 'maillog-(\d{8})', '$1'), "yyyyMMdd", $null))) {
        Write-Warning "No date found in file name: $logFile"
        continue
    }
    if ($date -lt $startDate -or $date -gt $endDate) {
        Write-Warning "Log file name out of range $startDate - $endDate"
        continue
    }
    if ($logFile.Extension -eq '.gz') {
        $fs = [System.IO.File]::OpenRead($logFile.FullName)
        $gz = New-Object System.IO.Compression.GzipStream $fs, ([IO.Compression.CompressionMode]::Decompress)
        $sr = New-Object System.IO.StreamReader $gz
    } else {
        $sr = [System.IO.StreamReader]::new($logFile.FullName)
    }
    $origHash = @{}
    $lineCount = 0
    $regex1 = [regex]"(?i)orig_to=<(?<name>\S+)@sloan.mit.edu>.+status=sent \(250"
    # While we are searching, might as well add up ALL of the mail received. Probably the fastest way to go.
    while ($line = $sr.ReadLine()) {
        if ($line -match $regex1) {
            $origHash[$Matches.name]++
        }
        $lineCount++
    }
    $totalLines += $lineCount
    if ($logFile.Extension -eq '.gz') {
        $fs.Dispose()
        $gz.Dispose()
    }
    $sr.Dispose()
    # Now make an aliases ordered hashtable by copying just the names and numbers we want.
    $aliasesHt = [ordered]@{}
    $aliasesHt["Date"] = $date.ToString("MM-dd-yyyy")
    $aliases.ForEach({
        $aliasesHt[$_] = $origHash[$_]
    })
    [PSCustomObject]$aliasesHt 
}

# Export results to CSV
$myCsv | Export-Csv -LiteralPath $outputCsv -NoTypeInformation -Encoding ASCII
Write-Verbose("{0:N0} total lines searched" -f $totalLines) -Verbose
Write-Host("Results in $outputCsv") -ForegroundColor Green
#eof