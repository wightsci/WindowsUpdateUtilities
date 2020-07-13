[CmdletBinding()]
Param()

If ($ReBootMayBeRequired -eq $True) {
    Write-Output "A reboot may be required by these updates."
}

$Install = Read-Host "Would you like to install these updates now? (Y/N)"

If ($Install.ToLower() -eq 'y') {
    Write-Output "Installing updates..."
    $UpdateInstaller = $UpdateSession.CreateUpdateInstaller()
    $UpdateInstaller.Updates = $UpdateInstallCollection
    $UpdateInstallResult = $UpdateInstaller.Install()

    Write-Host "Installation result: $($UpdateInstallResult.ResultCode)"
    Write-Host "Reboot required: $($UpdateInstallResult.RebootRequired)"
    Write-Host ""
    Write-Host "Detailed results:"

    For ($I = 0; $I -le $UpdateInstallCollection.Count -1; $I++){
        $Update = $UpdateInstallCollection.Item($I)
        Write-Output "$($I + 1)> $($Update.Title): $($UpdateInstallResult.GetUpdateResult($I).Resultcode)"
    }

}