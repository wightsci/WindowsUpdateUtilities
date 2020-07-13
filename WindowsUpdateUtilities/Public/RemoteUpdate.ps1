<#
.SYNOPSIS 

Scans a computer for Windows updates and installs them

.DESCRIPTION

Scans a computer for Windows updates and installs them, or creates a Scheduled Task to do so.

.PARAMETER AddTask
Specifies that a Scheduled Task should be created.

.PARAMETER Run
Specifies that a scan should be run interactively

.PARAMETER Format
Specifies the report format: CSV, HTML or XML for files, Console for the interactive screen.

.PARAMETER Path
Specifies the file path for the report file.

.PARAMETER StartAt
Specifies a Date/Time that the Scheduled Task should start. By default the task will start 45 seconds after it is created.

.PARAMETER ReportOnly
Specifies that a report is generated but no updates are downloaded or installed.

.INPUTS
None. You cannot pipe objects to this script.

.OUTPUTS
None to the pipeline.

.EXAMPLE
RemoteUpdate.ps1 -AddTask

This example creates a Scheduled Task that will run in 45 seconds time.

.EXAMPLE
RemoteUpdate.ps1 -Run

This example starts a scan and install.

.EXAMPLE
RemoteUpdate.ps1 -Run -ReportOnly

This example starts a scan and generates a report, but does not install any updates.

.NOTES
#>


Param(
    [Parameter(Mandatory=$true,ParameterSetName='Task')]
    [Switch]
    $AddTask,
    [Parameter(Mandatory=$false,ParameterSetName='Task')]
    [DateTime]
    $StartAt,
    [Parameter(Mandatory=$true,ParameterSetName="Exec")]
    [Switch]
    $Run,
    [Parameter(ParameterSetName="Exec")]
    [Parameter(ParameterSetName="Task")]
    [ValidateSet("CSV","XML","Console","HTML")]
    [String]
    $Format="csv",
    [Parameter(Mandatory=$false,ParameterSetName="Exec")]
    [Parameter(Mandatory=$false,ParameterSetName="Task")]
    [String]
    $Path,
    [Parameter(Mandatory=$false,ParameterSetName="Exec")]
    [Parameter(Mandatory=$false,ParameterSetName="Task")]
    [Switch]
    $ReportOnly
)

$ScriptGuid = 'b0c433cc-548c-4c8d-9f7d-d180effaf044'

## Constant Enums for Schedule Tasks. Derived from taskschd.h
Add-Type -TypeDefinition @" 
public enum TASK_RUN_FLAGS
    {
        TASK_RUN_NO_FLAGS	= 0,
        TASK_RUN_AS_SELF	= 0x1,
        TASK_RUN_IGNORE_CONSTRAINTS	= 0x2,
        TASK_RUN_USE_SESSION_ID	= 0x4,
        TASK_RUN_USER_SID	= 0x8
    }
public enum TASK_ENUM_FLAGS
    {
        TASK_ENUM_HIDDEN	= 0x1
    }
public enum TASK_LOGON_TYPE
    {
        TASK_LOGON_NONE	= 0,
        TASK_LOGON_PASSWORD	= 1,
        TASK_LOGON_S4U	= 2,
        TASK_LOGON_INTERACTIVE_TOKEN	= 3,
        TASK_LOGON_GROUP	= 4,
        TASK_LOGON_SERVICE_ACCOUNT	= 5,
        TASK_LOGON_INTERACTIVE_TOKEN_OR_PASSWORD	= 6
    }
public enum TASK_RUNLEVEL
    {
        TASK_RUNLEVEL_LUA	= 0,
        TASK_RUNLEVEL_HIGHEST	= 1
    }
public enum TASK_PROCESSTOKENSID
    {
        TASK_PROCESSTOKENSID_NONE	= 0,
        TASK_PROCESSTOKENSID_UNRESTRICTED	= 1,
        TASK_PROCESSTOKENSID_DEFAULT	= 2
    }
public enum TASK_STATE
    {
        TASK_STATE_UNKNOWN	= 0,
        TASK_STATE_DISABLED	= 1,
        TASK_STATE_QUEUED	= 2,
        TASK_STATE_READY	= 3,
        TASK_STATE_RUNNING	= 4
    }
public enum TASK_CREATION
    {
        TASK_VALIDATE_ONLY	= 0x1,
        TASK_CREATE	= 0x2,
        TASK_UPDATE	= 0x4,
        TASK_CREATE_OR_UPDATE	= ( TASK_CREATE | TASK_UPDATE ) ,
        TASK_DISABLE	= 0x8,
        TASK_DONT_ADD_PRINCIPAL_ACE	= 0x10,
        TASK_IGNORE_REGISTRATION_TRIGGERS	= 0x20
    }
public enum TASK_TRIGGER_TYPE2
    {
        TASK_TRIGGER_EVENT	= 0,
        TASK_TRIGGER_TIME	= 1,
        TASK_TRIGGER_DAILY	= 2,
        TASK_TRIGGER_WEEKLY	= 3,
        TASK_TRIGGER_MONTHLY	= 4,
        TASK_TRIGGER_MONTHLYDOW	= 5,
        TASK_TRIGGER_IDLE	= 6,
        TASK_TRIGGER_REGISTRATION	= 7,
        TASK_TRIGGER_BOOT	= 8,
        TASK_TRIGGER_LOGON	= 9,
        TASK_TRIGGER_SESSION_STATE_CHANGE	= 11,
        TASK_TRIGGER_CUSTOM_TRIGGER_01	= 12
    }
public enum TASK_SESSION_STATE_CHANGE_TYPE
    {
        TASK_CONSOLE_CONNECT	= 1,
        TASK_CONSOLE_DISCONNECT	= 2,
        TASK_REMOTE_CONNECT	= 3,
        TASK_REMOTE_DISCONNECT	= 4,
        TASK_SESSION_LOCK	= 7,
        TASK_SESSION_UNLOCK	= 8
    }
public enum TASK_ACTION_TYPE
    {
        TASK_ACTION_EXEC	= 0,
        TASK_ACTION_COM_HANDLER	= 5,
        TASK_ACTION_SEND_EMAIL	= 6,
        TASK_ACTION_SHOW_MESSAGE	= 7
    }
public enum TASK_INSTANCES_POLICY
    {
        TASK_INSTANCES_PARALLEL	= 0,
        TASK_INSTANCES_QUEUE	= 1,
        TASK_INSTANCES_IGNORE_NEW	= 2,
        TASK_INSTANCES_STOP_EXISTING	= 3
    }
public enum TASK_COMPATIBILITY
    {
        TASK_COMPATIBILITY_AT	= 0,
        TASK_COMPATIBILITY_V1	= 1,
        TASK_COMPATIBILITY_V2	= 2,
        TASK_COMPATIBILITY_V2_1	= 3,
        TASK_COMPATIBILITY_V2_2	= 4,
        TASK_COMPATIBILITY_V2_3	= 5,
        TASK_COMPATIBILITY_V2_4	= 6
    }
"@
#Constants from wuapi.h
Add-Type -TypeDefinition @"
public enum OperationResultCode
    {
        orcNotStarted	= 0,
        orcInProgress	= 1,
        orcSucceeded	= 2,
        orcSucceededWithErrors	= 3,
        orcFailed	= 4,
        orcAborted	= 5
    }
"@
#Constant from wuapicommon.h
Add-Type -TypeDefinition @"
public enum ServerSelection
    {
        ssDefault	= 0,
        ssManagedServer	= 1,
        ssWindowsUpdate	= 2,
        ssOthers	= 3
    } 
"@

if ($Run.IsPresent) {
#Need to declare these all here because they don't seem to survive being passed as parameters very well
$UpdateSession = New-Object -ComObject "Microsoft.Update.Session"
$UpdateCollection =  New-Object -ComObject "Microsoft.Update.UpdateColl"
$UpdateInstallCollection = New-Object -ComObject "Microsoft.Update.UpdateColl"
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$UpdateDownloader = $UpdateSession.CreateUpdateDownloader()
$UpdateInstaller = $UpdateSession.CreateUpdateInstaller()
}

$WorkDirectory = 'C:\UpdateScan'

if (!($Path)) {
    $Path = "C:\UpdateScan\UpdateScan.$($Format)"
}

Function Remove-UpdateTask {
    $STService = New-Object -ComObject Schedule.Service 
    $STService.Connect()
    $RootFolder = $STService.GetFolder("\")
    try {
        $RootFolder.DeleteTask($Script:ScriptGuid,$Null)
        Write-Verbose "Remove-UpdateTask: Task $ScriptGuid Removed."
    }
    catch {}
}
Function Add-UpdateTask {
    $STService = New-Object -ComObject Schedule.Service 
    $STService.Connect()
    $RootFolder = $STService.GetFolder("\")

    $NewTaskDef = $STService.NewTask(0)
    $RegInfo = $NewTaskDef.RegistrationInfo
    $RegInfo.Description = "Update Scan and Install"
    $RegInfo.Author = "Stuart Squibb"

    $Principal = $NewTaskDef.Principal
    $Principal.LogonType = [TASK_LOGON_TYPE]::Task_Logon_Service_Account
    $Principal.UserId = 'NT AUTHORITY\SYSTEM'
    $Principal.Id = "System"
    $Principal | Select-Object * | Write-Verbose
    $Settings = $NewTaskDef.Settings
    $Settings.Enabled = $True
    $Settings.DisallowStartIfOnBatteries = $False

    $Trigger = $NewTaskDef.Triggers.Create([TASK_TRIGGER_TYPE2]::TASK_TRIGGER_TIME)

    if ($Script:StartAt) {
        $StartTime = $Script:StartAt
        Write-Verbose $StartTime
    }
    else {
        $StartTime = (Get-Date).AddSeconds(45)
    }
     
    $EndTime = ($StartTime.AddMinutes(5)).ToString("yyyy-MM-ddTHH:mm:ss")
    $StartTime = $StartTime.toString("yyyy-MM-ddTHH:mm:ss")

    Write-Verbose "Add-UpdateTask: Time Now  : $((Get-Date).ToString('yyyy-MM-ddTHH:mm:ss'))"
    Write-Verbose "Add-UpdateTask: Start Time: $($StartTime)"
    Write-Verbose "Add-UpdateTask: End Time  : $($EndTime)"

    $Trigger.StartBoundary = $StartTime
    $Trigger.EndBoundary = $EndTime
    $Trigger.ExecutionTimeLimit = "PT60M"
    $Trigger.Id = "TimeTriggerId"
    $Trigger.Enabled = $True 

    $Action = $NewTaskDef.Actions.Create([TASK_ACTION_TYPE]::TASK_ACTION_EXEC)
    $Action.Path = "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"
    if ($ReportOnly.IsPresent) {
        $Action.Arguments = "-ExecutionPolicy ByPass -NoProfile -NonInteractive -File C:\UpdateScan\RemoteUpdate.ps1 -Run -Format $Format -Path $Path -ReportOnly"
    }
    Else {
        $Action.Arguments = "-ExecutionPolicy ByPass -NoProfile -NonInteractive -File C:\UpdateScan\RemoteUpdate.ps1 -Run -Format $Format -Path $Path"
    }
    
    $Action.WorkingDirectory = "C:\UpdateScan"

    Write-Verbose "Add-UpdateTask: Task Definition created. About to submit Task..."

    [Void]$RootFolder.RegisterTaskDefinition($ScriptGuid, $NewTaskDef,[TASK_CREATION]::TASK_CREATE_OR_UPDATE,$Null,$Null,$Null)

    Write-Verbose "Add-UpdateTask: Task $ScriptGuid Submitted."
}



# get-updatecollection
Function Get-UpdateCollection {
    Write-Verbose "Get-UpdateCollection: Searching for updates..."
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and BrowseOnly=0")
    $Updates = $SearchResult.Updates
    
    Write-Verbose "Get-UpdateCollection: $($Updates.Count) updates found."

    If ($Updates.Count -eq 0) {
        #This area blank by design
     }
    else {  
        return $SearchResult    
    }
}

function Get-DownloadCollection {
    param(
        $SearchResult
    )
    Write-Verbose "Get-DownloadCollection: $($SearchResult.Updates.Count) updates"
    For ($I = 0; $I -le $SearchResult.Updates.Count-1; $I++) {
        $Update = $SearchResult.Updates.Item($I)
        $AddThisUpdate = $false
        If ($Update.InstallationBehavior.CanRequestUserInput -eq $True) {
            Write-Verbose "$($I + 1)> Skipping $($Update.Title) because of required user input"
        }
        Else {
            If ($Update.EulaAccepted -eq $False) {
                Write-Verbose "$($I + 1)> $($Update.Title) has a licence agreement:"
                Write-Verbose $Update.EulaText
                $Update.AcceptEula()
                Write-Verbose "$($I + 1)> EULA Accepted"
                $AddThisUpdate = $True
            }
            Else {
                $AddThisUpdate = $True
            }
        }
        If ($AddThisUpdate -eq $True) {
            Write-Verbose "$($I + 1)> Adding: $($Update.Title)"
            [void]$UpdateCollection.Add($Update)
        }
    }
    If ($UpdateCollection.Count -eq 0) {
        Write-Verbose "Get-DownloadCollection: All applicable updates were skipped."
        Exit
    } 
    else {
        Write-Verbose "Get-DownloadCollection: $($UpdateCollection.Count) updates available to download."
    }
}

Function Get-Updates {
    Write-Verbose "Get-Updates: $($UpdateCollection.Count) updates to download."
    $UpdateDownloader.Updates = $UpdateCollection

    Write-Verbose "Get-Updates: Downloading updates..."
    [void]$UpdateDownloader.Download()

    $ReBootMayBeRequired = $false

    Write-Verbose "Get-Updates: Download complete for:"

    For ($I = 0; $I -le $UpdateCollection.Count-1; $I++) {
        $Update = $UpdateCollection.Item($I)
        If ($Update.IsDownloaded -eq $True) {
            Write-Verbose "$($I + 1)> $($Update.Title)"
            [void]$UpdateInstallCollection.Add($Update)
            If ($Update.InstallationBehavior.RebootBehavior -gt 0) {
                $ReBootMayBeRequired = $True
            }
        }
    }

    If ($UpdateInstallCollection.Count  -eq 0) {
        Write-Verbose "Get-Updates: No updates were able to be downloaded."
        Exit
    }
    else {
        Write-Verbose ("Get-Updates: {0} updates downloaded ready to install." -f $UpdateCollection.Count)
        If ($ReBootMayBeRequired -eq $True) {
            Write-Verbose "A reboot may be required by these updates."
        }  
    }
}

Function Install-Updates {
    if ($UpdateInstallCollection.Count -ne 0) {
        Write-Verbose ("Install-Updates: Installing {0} updates..." -f $UpdateInstallCollection.Count)
        $UpdateInstaller.Updates = $UpdateInstallCollection
        $UpdateInstallResult = $UpdateInstaller.Install()

        Write-Verbose "Install-Updates: Installation result: $($UpdateInstallResult.ResultCode)"
        Write-Verbose "Install-Updates: Reboot required: $($UpdateInstallResult.RebootRequired)"
        Write-Verbose "Install-Updates: "
        Write-Verbose "Install-Updates: Detailed results:"

        For ($I = 0; $I -le $UpdateInstallCollection.Count -1; $I++){
            $Update = $UpdateInstallCollection.Item($I)
            Write-Verbose "$($I + 1)> $($Update.Title): $($UpdateInstallResult.GetUpdateResult($I).Resultcode)"
        }
    }
    else {
        Write-Verbose "Install-Updates: No updates to install."
    }
}

Function Export-InstalledUpdateCollection {
    Param (
        [Parameter(Mandatory=$True)]
        [ValidateSet("xml","csv","console","html")]
        [String]
        $Format,
        [Parameter(Mandatory=$False)]
        [String]
        $FileName
    )
    $UpdateObjCol = @()
    For ($I = 0; $I -le $UpdateInstallCollection.Count -1; $I++){
        $Update = $UpdateInstallCollection.Item($I)
        $UpdateObj = [PSCustomObject]@{
            Title = $Update.Title
            ResultCode = $UpdateInstallCollection.InstallationResult.GetUpdateResult($I).Resultcode
        }
        $UpdateObjCol += $UpdateObj
    }

    $OutPutObject = $UpdateObjCol  
    switch ($Format) {
            'csv'  { $OutPutObject | Export-Csv -Path $FileName -NoTypeInformation  }
            'xml'  { ($OutPutObject | ConvertTo-Xml -NoTypeInformation -As Document).OuterXML | Out-File -FilePath $FileName }
            'html' { $OutPutObject | ConvertTo-Html -Title "Installed Windows Updates for $env:ComputerName" | Out-File -FilePath $FileName}
            'console' { Format-Table -InputObject $OutPutObject}
        }
    }

# export-requiredupdatecollection - reports on needed updates
Function Export-RequiredUpdateCollection {
Param (
    [Parameter(Mandatory=$True)]
    [ValidateSet("xml","csv","console","html")]
    [String]
    $Format,
    [Parameter(Mandatory=$False)]
    [String]
    $FileName,
    [Parameter(ValueFromPipeline=$True,Mandatory=$True)]
    [Object]
    $UpdateCollection
)
$OutPutObject = $UpdateCollection.Updates | Select-Object -Property MsrcSeverity, Title, MaxDownloadSize, MinDownloadSize, @{Name="KBs";Expression={$_.KBArticleIds -join ';'}}    
switch ($Format) {
        'csv'  { $OutPutObject | Export-Csv -Path $FileName -NoTypeInformation  }
        'xml'  { ($OutPutObject | ConvertTo-Xml -NoTypeInformation -As Document).OuterXML | Out-File -FilePath $FileName }
        'html' { $OutPutObject | ConvertTo-Html -Title "Needed Windows Updates for $env:ComputerName" | Out-File -FilePath $FileName}
        'console' { Format-Table -InputObject $OutPutObject}
    }
}

    


if ($AddTask.IsPresent) {
    Add-UpdateTask
}

if ($Run.IsPresent) {
    try {
        New-Item -ItemType Directory -Path $WorkDirectory -ErrorAction Stop
    }
    catch {}

    try {
        # Copy-Item $CabSource -Destination $CabLocation -ErrorAction Stop
    }
    catch {}

    Write-Verbose "Exporting $Format format file to $Path"
    $ExportCollection = Get-UpdateCollection
    If ($ExportCollection) {
        Write-Verbose "Export Collection has $($ExportCollection.Updates.Count) updates"
        Export-RequiredUpdateCollection -Format $Format -FileName $Path -UpdateCollection $ExportCollection
        if (!$ReportOnly.IsPresent) {
            Get-DownloadCollection -SearchResult $ExportCollection
            Write-Verbose "Download Collection has $($UpdateCollection.Count) updates."
            Get-Updates
            Install-Updates
            Export-InstalledUpdateCollection -FileName "$($WorkDirectory)\InstallationReport.html" -Format html
        }    
        Remove-Updatetask
    }
    Else {
        Write-Verbose "No updates to process."
    }
}





