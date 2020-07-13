Function Get-WindowsUpdateList{
[CmdletBinding()]
Param()

$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()

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
}
