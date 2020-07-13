[CmdletBinding()]
Param()

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

# Loader for external modules
$ScriptRoot = Split-Path $Script:MyInvocation.MyCommand.Path
$Public = @( Get-ChildItem -Path $ScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue )
$Private = @( Get-ChildItem -Path $ScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue )
@($Public + $Private) | Foreach-Object { . $_.FullName }
Export-ModuleMember -Function $Public.Basename