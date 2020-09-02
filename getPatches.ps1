Import-Module pswindowsupdate

$patchCat = "Office 2016" #replace "Office 2016" with $args[0] to be able to pass the patch category as script parameter

$getPatches = Get-WUList -Category $patchCat
$formatedPatchlist = $getPatches | Out-String

$logFileExists = [System.Diagnostics.EventLog]::SourceExists("Patch Check");
#if the log source exists, it uses it. if it doesn't, it creates it.
if (-not $logFileExists)
{New-EventLog -LogName Application -Source "Patch Check"} 
#writes events to the windows event log

$eventlogvalues= @{
LogName = 'Application' 
Source = 'Patch Check'
}

$errorEntry = @{
EntryType = "Error"
EventID = 818
Message = "There were no patches found matching the specified category:`n $patchCat "
}

$infoEntry = @{
EntryType = "Information"
EventID = 816
Message = "We found the following matching patches:`n$formatedPatchList"
}


if ($getPatches -ne $null)
{
Write-EventLog @eventlogvalues @infoentry
#uncomment to also install, if any have been found. Event will still be triggered, but can be modified.
#Get-WUList -ForceInstall
}
else 
{
Write-EventLog @eventlogvalues @errorentry
}