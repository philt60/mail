Set-StrictMode -Version Latest

Set-Location $PSScriptRoot
$badList = Get-Content("badList.txt")
foreach ($afile in ("aliases", "alias_clb_off")) {
    $outFile = $($afile + "2.txt")
    Write-Verbose "outFile: $outFile" -Verbose
    Remove-Item -Path $outFile -ErrorAction SilentlyContinue
    # Get the aliases, skipping comments (#) - with a negative look ahead.
    Get-Content -Path $afile -ea Stop | % {
        if ($_ -match '^(?!#)(?<name>\S+)\s*:\s*(?<list>\S+.*)$') {
            if ($Matches.name -in $badList) {
                Write-Host "Removing $_" -ForegroundColor Yellow
            } else {
                $_ | Tee-Object -FilePath $outFile -Append
            }
        } else {
            $_ | Tee-Object -FilePath $outFile -Append
        }
    }
}


