Function Download-WindowsUpdates {
[CmdletBinding()]
Param()

    Write-Output "Creating download list..."

    $UpdateCollection =  New-Object -ComObject "Microsoft.Update.UpdateColl"

    For ($I = 0; $I -le $SearchResult.Updates.Count-1; $I++) {
        $Update = $SearchResult.Updates.Item($I)
        $AddThisUpdate = $false
        If ($Update.InstallationBehavior.CanRequestUserInput -eq $True) {
            Write-OutPut "$($I + 1)> Skipping $($Update.Title) because of required user input"
        }
        Else {
            If ($Update.EulaAccepted -eq $False) {
                Write-OutPut "$($I + 1)> $($Update.Title) has a licence agreement:"
                Write-Output $Update.EulaText
                $Acceptance = Read-Host "Do you accept the agreement? (Y/N)" 
                if ($Acceptance.ToLower() -eq 'y') {
                    $Update.AcceptEula()
                    $AddThisUpdate = $True
                }
                Else {
                    Write-OutPut "$($I + 1)> Skipping $($Update.Title) because of decline licence agreement."
                }
            }
            Else {
                $AddThisUpdate = $True
            }
        }
        If ($AddThisUpdate -eq $True) {
            Write-OutPut "$($I + 1)> Adding: $($Update.Title)"
            [void]$UpdateCollection.Add($Update)
        }

    }

    If ($UpdateCollection.Count -eq 0) {
        Write-OutPut "All applicable update were skipped."
    }
    Else {

        Write-Output "Downloading updates..."

        $UpdateDownloader = $UpdateSession.CreateUpdateDownloader()
        
        $UpdateDownloader.Updates = $UpdateCollection
        
        [void]$UpdateDownloader.Download()
        
        $UpdateInstallCollection = New-Object -ComObject "Microsoft.Update.UpdateColl"
        
        $ReBootMayBeRequired = $false
        
        Write-Output "Download complete for:"
        
        For ($I = 0; $I -le $SearchResult.Updates.Count-1; $I++) {
            $Update = $SearchResult.Updates($I)
            If ($Update.IsDownloaded -eq $True) {
                Write-OutPut "$($I + 1)> $($Update.Title)"
                [void]$UpdateInstallCollection.Add($Update)
                If ($Update.InstallationBehavior.RebootBehavior -gt 0) {
                    $ReBootMayBeRequired = $True
                }
            }
        }
    }

    If ($UpdateInstallCollection.Count  -eq 0) {
        Write-Output "No updates were able to be downloaded."
    }
}