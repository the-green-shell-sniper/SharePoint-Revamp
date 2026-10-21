#requires -Version 7.4.6
<#
.SYNOPSIS
    Queue checklist sites in at most four elevated, retained PowerShell windows.
.DESCRIPTION
    Explicit -Apply starts Match -> auto approve scores >= 0.80 -> DryRun -> Apply
    for each web, parent before children. Without -Apply only lists the selected queue.
    Low-score matches remain unapproved. Failed preflight stops that site's worker.
    One worker owns a checklist site and sequential descendants; target-only webs
    are still recursively reported. Related site collections never run concurrently.
    Completed windows stay open but release their active worker slot via status files.
    Do not concurrently run a separate manual permission/scout process on these sites.
#>
[CmdletBinding()]
param(
    [string]$ChecklistFile = (Join-Path $PSScriptRoot '..\ReferenceFiles\NewSiteChecklist.csv'),
    [string]$OutputFolder = (Join-Path $PSScriptRoot '..\Reports\PermRuns'),
    [string]$ReviewRoot = (Join-Path $PSScriptRoot '..\Reports\PermissionsReports\PermissionReviews'),
    [string[]]$TargetSites,
    [ValidateRange(1, 4)][int]$MaxConcurrent = 4,
    [string]$ClientId = '2e20d82f-56ba-418e-b465-9718493f0df8',
    [string]$Tenant,
    [switch]$NoContent,
    [switch]$SkipPersistLoginSetup,
    [switch]$Apply,
    [string]$WorkerJobFile
)
. (Join-Path $PSScriptRoot 'SPOutputLimits.Common.ps1')
. (Join-Path $PSScriptRoot 'Invoke-SPPermissionChecklist.ps1') -ChecklistFile $ChecklistFile `
    -OutputFolder $OutputFolder -ReviewRoot $ReviewRoot -TargetSites $TargetSites `
    -ClientId $ClientId -Tenant $Tenant -NoContent:$NoContent -SkipPersistLoginSetup:$SkipPersistLoginSetup

function Get-SPQueueCollectionKey {
    param([string]$Url)
    $uri = [uri]$Url
    $segments = @($uri.AbsolutePath.Trim('/') -split '/')
    $path = if ($segments.Count -ge 2 -and $segments[0] -in @('sites', 'teams')) {
        '/' + ($segments[0..1] -join '/')
    } else { '/' }
    ($uri.GetLeftPart([UriPartial]::Authority) + $path).ToLowerInvariant()
}

function Test-SPQueueConflict {
    param($Candidate, [object[]]$Active)
    foreach ($job in $Active) {
        if ((Get-SPQueueCollectionKey $Candidate.TargetUrl) -eq (Get-SPQueueCollectionKey $job.TargetUrl)) { return $true }
    }
    return $false
}

function Invoke-SPPermissionQueueWorker {
    param([string]$Path)
    $job = Import-Clixml -LiteralPath $Path -ErrorAction Stop
    $statusPath = Join-Path $job.OutputFolder 'WorkerStatus.json'
    $status = [ordered]@{
        TargetUrl = $job.TargetUrl; Status = 'Running'; Error = ''; StartedUtc = [datetime]::UtcNow.ToString('o')
        Transcript = 'Disabled to keep generated files below the repository size limit; inspect structured evidence and console output.'
    }
    $status | ConvertTo-Json | Set-Content -LiteralPath "$statusPath.tmp" -Encoding utf8
    Move-Item -LiteralPath "$statusPath.tmp" -Destination $statusPath -Force
    $transcribing = $false
    $slots = $null; $collectionLock = $null
    $hasSlot = $false; $hasCollectionLock = $false
    try {
        $collectionKey = Get-SPQueueCollectionKey $job.TargetUrl
        $collectionHash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData(
            [Text.Encoding]::UTF8.GetBytes($collectionKey)))
        $collectionLock = [Threading.Mutex]::new($false, "Global\SharePointRevamp-$collectionHash")
        try { $hasCollectionLock = $collectionLock.WaitOne() }
        catch [Threading.AbandonedMutexException] {
            $hasCollectionLock = $true
            Write-Warning 'An earlier collection worker ended unexpectedly; live preflight will re-read its state.'
        }
        # Survives a scheduler restart while any workers remain alive.
        while (-not $hasSlot) {
            foreach ($slot in 1..4) {
                $candidateLock = [Threading.Mutex]::new($false, "Global\SharePointRevampPermissionSlot-$slot")
                try { $hasSlot = $candidateLock.WaitOne(0) }
                catch [Threading.AbandonedMutexException] {
                    $hasSlot = $true
                    Write-Warning 'An earlier worker abandoned its active slot; recovering it.'
                }
                if ($hasSlot) { $slots = $candidateLock; break }
                $candidateLock.Dispose()
            }
            if (-not $hasSlot) { Start-Sleep -Seconds 1 }
        }
        $options = @{
            ChecklistFile = $job.ChecklistFile; TargetSites = @($job.TargetUrl)
            Mode = 'Auto'; OutputFolder = $job.OutputFolder; ReviewRoot = $job.ReviewRoot
            ClientId = $job.ClientId; Tenant = $job.Tenant
            NoContent = $job.NoContent; SkipPersistLoginSetup = $job.SkipPersistLoginSetup
        }
        & (Join-Path $PSScriptRoot 'Invoke-SPPermissionChecklist.ps1') @options
        $pending = @(
            foreach ($file in Get-ChildItem -LiteralPath $job.ReviewRoot -Recurse -Filter '*Review.csv') {
                Import-Csv -LiteralPath $file.FullName | Where-Object { $_.Approved -notmatch '^(?i:Yes|True|1|Y)$' }
            }
        )
        $status.Status = if ($pending.Count) { 'Completed - manual review pending' } else { 'Completed' }
        $status['PendingReviewRows'] = $pending.Count
    } catch {
        $status.Status = 'Failed'; $status.Error = $_.Exception.Message
        Write-Warning "Site worker FAILED: $($_.Exception.Message). Earlier applied changes remain; inspect Changes.csv and per-web logs."
    } finally {
        if ($hasSlot) { $slots.ReleaseMutex() }
        if ($slots) { $slots.Dispose() }
        if ($hasCollectionLock) { $collectionLock.ReleaseMutex() }
        if ($collectionLock) { $collectionLock.Dispose() }
        try {
            $status['FinishedUtc'] = [datetime]::UtcNow.ToString('o')
            # Publish completion atomically; the scheduler must not consume partial JSON.
            $temporary = "$statusPath.tmp"
            $status | ConvertTo-Json | Set-Content -LiteralPath $temporary -Encoding utf8
            Move-Item -LiteralPath $temporary -Destination $statusPath -Force
        } catch { $status.Status = 'Failed'; $status.Error += "; Could not publish worker status: $($_.Exception.Message)"; throw }
        Write-Host "[$($status.Status)] $($job.TargetUrl) - reports: $($job.OutputFolder)"
        Write-Host 'This window is retained for inspection. The queue may now start another site.'
    }
}

function Invoke-SPPermissionQueue {
    $ErrorActionPreference = 'Stop'
    $pairs = @(Get-SPChecklistPairs $ChecklistFile)
    if ($TargetSites.Count) {
        foreach ($url in $TargetSites) {
            if ($url.TrimEnd('/') -inotin $pairs.TargetUrl) { throw "Unknown checklist target: $url" }
        }
        $selected = @($TargetSites | ForEach-Object { $_.TrimEnd('/') })
        $pairs = @($pairs | Where-Object { $_.TargetUrl -iin $selected })
    }
    if (-not $Apply) {
        $pairs | Format-Table SourceUrl, TargetUrl -AutoSize
        Write-Host "Preview only. Use -Apply to authorize automatic Match/DryRun/Apply for $($pairs.Count) queued sites (maximum $MaxConcurrent active)."
        return
    }
    if (-not $IsWindows) { throw 'Elevated retained workers require Windows PowerShell 7.' }
    $mutex = [Threading.Mutex]::new($false, 'Global\SharePointRevampPermissionQueue')
    $ownsMutex = $false
    try {
        try { $ownsMutex = $mutex.WaitOne(0) } catch [Threading.AbandonedMutexException] { $ownsMutex = $true }
        if (-not $ownsMutex) { throw 'Another permission queue is already active. Use its queue status instead of starting another.' }
        $root = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputFolder)
        $reviews = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ReviewRoot)
        $checklistPath = (Resolve-Path -LiteralPath $ChecklistFile -ErrorAction Stop).ProviderPath
        $batch = Join-Path $root ('q-' + (Get-Date -Format 'yyMMddHHmmssfff'))
        [void](Assert-SPOutputPath $batch)
        [void][IO.Directory]::CreateDirectory($batch)
        $pending = [Collections.Generic.List[object]]::new()
        foreach ($pair in $pairs) { $pending.Add($pair) }
        $active = [Collections.Generic.List[object]]::new()
        $results = [Collections.Generic.List[object]]::new()
        $workerScript = Join-Path $PSScriptRoot 'Invoke-SPPermissionQueue.ps1'
        $executable = Join-Path $PSHOME 'pwsh.exe'
        while ($pending.Count -or $active.Count) {
            foreach ($entry in @($active.ToArray())) {
                $statusPath = Join-Path $entry.Folder 'WorkerStatus.json'
                $status = if (Test-Path -LiteralPath $statusPath) { Get-Content -LiteralPath $statusPath -Raw | ConvertFrom-Json } else { $null }
                if ($status -and $status.Status -ne 'Running') {
                    $results.Add($status)
                    [void]$active.Remove($entry)
                } elseif ($entry.Process.HasExited) {
                    $results.Add([pscustomobject]@{ TargetUrl = $entry.TargetUrl; Status = 'Failed'; Error = 'Window exited before publishing completion; inspect partial reports.' })
                    [void]$active.Remove($entry)
                }
            }
            for ($index = 0; $index -lt $pending.Count -and $active.Count -lt $MaxConcurrent;) {
                $pair = $pending[$index]
                if (Test-SPQueueConflict $pair $active.ToArray()) { $index++; continue }
                $key = (Get-SPChecklistSiteKey $pair.TargetUrl).Split('-')[-1]
                $folder = Join-Path $batch $key
                [void](Assert-SPOutputPath $folder)
                [void][IO.Directory]::CreateDirectory($folder)
                $jobPath = Join-Path $folder 'WorkerJob.clixml'
                @{
                    ChecklistFile = $checklistPath; SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl
                    OutputFolder = $folder; ReviewRoot = (Join-Path $reviews (Get-SPChecklistSiteKey $pair.TargetUrl))
                    ClientId = $ClientId; Tenant = $Tenant; NoContent = [bool]$NoContent
                    SkipPersistLoginSetup = [bool]$SkipPersistLoginSetup
                } | Export-Clixml -LiteralPath $jobPath
                $command = "& '$($workerScript.Replace("'", "''"))' -WorkerJobFile '$($jobPath.Replace("'", "''"))'"
                $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
                try {
                    $process = Start-Process -FilePath $executable -Verb RunAs -ArgumentList @('-NoLogo', '-NoProfile', '-NoExit', '-EncodedCommand', $encoded) -PassThru -ErrorAction Stop
                    $active.Add([pscustomobject]@{ SourceUrl = $pair.SourceUrl; TargetUrl = $pair.TargetUrl; Folder = $folder; Process = $process })
                } catch {
                    $results.Add([pscustomobject]@{ TargetUrl = $pair.TargetUrl; Status = 'Failed'; Error = "Worker launch failed: $($_.Exception.Message)" })
                    Write-Warning "Could not launch $($pair.TargetUrl): $($_.Exception.Message)"
                }
                $pending.RemoveAt($index)
            }
            @(
                $results
                foreach ($entry in $active) { [pscustomobject]@{ TargetUrl = $entry.TargetUrl; Status = 'Running'; Error = '' } }
                foreach ($entry in $pending) { [pscustomobject]@{ TargetUrl = $entry.TargetUrl; Status = 'Queued'; Error = '' } }
            ) | Select-Object TargetUrl, Status, Error | Export-Csv -LiteralPath (Join-Path $batch 'QueueStatus.csv') -NoTypeInformation -Encoding utf8
            if ($active.Count) { Start-Sleep -Seconds 3 }
        }
        Write-Host "Queue reports: $batch"
        if (@($results | Where-Object Status -eq 'Failed').Count) { throw 'One or more site workers failed. Inspect QueueStatus.csv and retained windows. Do not blindly repeat Apply.' }
    } finally {
        if ($ownsMutex) { $mutex.ReleaseMutex() }
        $mutex.Dispose()
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    if ($WorkerJobFile) { Invoke-SPPermissionQueueWorker $WorkerJobFile }
    else { Invoke-SPPermissionQueue }
}
