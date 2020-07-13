#Enums from wuapi.h
enum AutomaticUpdatesNotificationLevel
    {
        aunlNotConfigured	= 0
        aunlDisabled	= 1
        aunlNotifyBeforeDownload	= 2
        aunlNotifyBeforeInstallation	= 3
        aunlScheduledInstallation	= 4
    } 

enum AutomaticUpdatesScheduledInstallationDay
    {
        ausidEveryDay	= 0
        ausidEverySunday	= 1
        ausidEveryMonday	= 2
        ausidEveryTuesday	= 3
        ausidEveryWednesday	= 4
        ausidEveryThursday	= 5
        ausidEveryFriday	= 6
        ausidEverySaturday	= 7
    } 

enum DownloadPhase
    {
        dphInitializing	= 1
        dphDownloading	= 2
        dphVerifying	= 3
    } 

enum DownloadPriority
    {
        dpLow	= 1
        dpNormal	= 2
        dpHigh	= 3
        dpExtraHigh	= 4
    } 

enum AutoSelectionMode
    {
        asLetWindowsUpdateDecide	= 0
        asAutoSelectIfDownloaded	= 1
        asNeverAutoSelect	= 2
        asAlwaysAutoSelect	= 3
    } 

enum AutoDownloadMode
    {
        adLetWindowsUpdateDecide	= 0
        adNeverAutoDownload	= 1
        adAlwaysAutoDownload	= 2
    } 

enum InstallationImpact
    {
        iiNormal	= 0
        iiMinor	= 1
        iiRequiresExclusiveHandling	= 2
    } 

enum InstallationRebootBehavior
    {
        irbNeverReboots	= 0
        irbAlwaysRequiresReboot	= 1
        irbCanRequestReboot	= 2
    } 	

enum OperationResultCode
    {
        orcNotStarted	= 0
        orcInProgress	= 1
        orcSucceeded	= 2
        orcSucceededWithErrors	= 3
        orcFailed	= 4
        orcAborted	= 5
    } 	

enum UpdateType
    {
        utSoftware	= 1
        utDriver	= 2
    } 	

enum UpdateOperation
    {
        uoInstallation	= 1
        uoUninstallation	= 2
    } 

enum DeploymentAction
    {
        daNone	= 0
        daInstallation	= 1
        daUninstallation	= 2
        daDetection	= 3
    }

enum UpdateExceptionContext
    {
        uecGeneral	= 1
        uecWindowsDriver	= 2
        uecWindowsInstaller	= 3
    } 

enum AutomaticUpdatesUserType
    {
        auutCurrentUser	= 1
        auutLocalAdministrator	= 2
    } 	

enum AutomaticUpdatesPermissionType
    {
        auptSetNotificationLevel	= 1
        auptDisableAutomaticUpdates	= 2
        auptSetIncludeRecommendedUpdates	= 3
        auptSetFeaturedUpdatesEnabled	= 4
        auptSetNonAdministratorsElevated	= 5
    } 	

enum UpdateServiceRegistrationState
    {
        usrsNotRegistered	= 1
        usrsRegistrationPending	= 2
        usrsRegistered	= 3
    } 	

enum SearchScope
    {
        searchScopeDefault	= 0
        searchScopeMachineOnly	= 1
        searchScopeCurrentUserOnly	= 2
        searchScopeMachineAndCurrentUser	= 3
        searchScopeMachineAndAllUsers	= 4
        searchScopeAllUsers	= 5
    } 	

$UpdateSession = New-Object -ComObject "Microsoft.Update.Session"
$UpdateSession.ClientApplicationID = 'PowerShell WindowsUpdateUtilities'
$UpdateServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

Write-Output "Commencing Search for Updates..."

$SearchResult = $UpdateSearcher.Search("IsInstalled=0")

if ($SearchResult.Updates.Count -eq 0) {
    Write-Output "There are no applicable updates."
    Exit
}

Write-Output "List of applicable updates on the machine:"

For ($I = 0; $I -le $SearchResult.Updates.Count-1; $I++) {
    $Update = $SearchResult.Updates.Item($I)
    Write-Output "$($I + 1)> $($Update.Title)"
}

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
    Exit
}

Write-Output "Dowloading updates..."

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

If ($UpdateInstallCollection.Count  -eq 0) {
    Write-Output "No updates were able to be downloaded."
    Exit
}


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