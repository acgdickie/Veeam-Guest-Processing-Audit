#Requires -Version 7.0
<#
.SYNOPSIS
    Lists every VM that Veeam Backup & Replication v13 processes with guest processing
    (application-aware processing), with the job and the SQL transaction log settings.

.DESCRIPTION
    Checks every VM backup job, replication job and VM copy job (VMware vSphere and
    Microsoft Hyper-V), and every Veeam Agent job.

    A VM is reported when:
      - application-aware processing is enabled in the job, and
      - the VM is not excluded (job exclusions and global VM exclusions), and
      - application processing is not disabled for that VM (job > Guest Processing > Applications...).

    The script finds the VMs of each job from:
      - VMs added to the job directly
      - VMs inside containers added to the job (vCenter, datacenters, clusters, hosts,
        folders, resource pools, vApps, datastores, tags, Hyper-V hosts, clusters and
        VM groups), found from the virtual infrastructure inventory
      - the tasks of the latest job run (the latest full session and any retry,
        Quick Backup or planned failover sessions after it)

    The guest processing settings of a VM come from the deepest job object that holds
    the VM, the same way Veeam picks them:
      - a VM entry (in "Applications..." or added to the job) before any container
      - VMware: a tag before a VM folder, a VM folder before a datacenter, cluster, host,
        resource pool or datastore, and a deeper container before its parent
      - Hyper-V: a tag, then a VM group, a host, a cluster, a host group, SCVMM
    This is the Veeam default order. Two registry values on the backup server change it:
    TagsPriorityOverContainers = 0 and PrioritizeSpecialGuestProcessingCreds = 1 (the script
    warns when it runs on that server).
    Veeam has no job-level default for these settings. When the script cannot match a
    VM to a job object, it shows "(assumed)" when every possible object has the same
    settings, or Unknown when they differ. Always read the SettingsFrom column.

    SQL transaction log values in the SqlLogHandling column:
      Truncate logs | Do not truncate logs | Back up logs (log shipping) | Copy only
      N/A (Linux guest) - SQL Server processing runs on Windows guests only
      (Veeam Agent jobs can also show N/A (Mac guest) or N/A (Unix guest))

    Not checked: CDP policies (the script warns about each one that has application-aware
    processing), backup copy jobs (no guest processing), Veeam Agent jobs of the Workstation
    type (no application-aware processing), and other job types (the script lists them).
    For VMware Cloud Director jobs the VMs come from the job sessions only (the script warns).

    The result is shown on screen and saved to a CSV file. The script only reads data.
    It does not change any job.

.PARAMETER Server
    Name or IP address of the Veeam backup server. Default: localhost.
    If this PowerShell window is already connected to a backup server, the script uses that
    session and keeps it. The script stops when -Server names another server (or the same
    server by another name) or when -Credential is given: it does not end a session that it
    did not open.

.PARAMETER Credential
    Credentials for the backup server (DOMAIN\User or UPN). If you do not supply them,
    the current Windows account is used.

.PARAMETER Port
    Port of the backup server. If you do not supply it, the Veeam default is used.

.PARAMETER ForceAcceptTlsCertificate
    Accepts the TLS certificate of the backup server without a prompt.

.PARAMETER JobName
    One or more job names to check. An exact job name is matched as typed, also when it
    contains [ or ]. Otherwise wildcards (* ? [ ]) apply. Put a backtick before [ or ]
    to match it literally in a wildcard pattern. Default: all jobs.

.PARAMETER OutputPath
    Path of the CSV file, or an existing folder for a time-stamped file. Default: a
    time-stamped file in the script folder. If the file cannot be written, the script
    saves it to the TEMP folder and tells you where.

.PARAMETER IncludeDisabled
    Also shows VMs where guest processing is off (in the job, or for the VM).

.PARAMETER DetectSql
    Adds Yes/No to the SqlFound column. Yes means that a SQL Server application restore
    point exists for a machine with the same name (in any job, in any restore point that
    is still kept). The match is by name only: VMs with the same name share one result,
    a VM renamed since its last backup shows No until its next backup, and VMs that were
    never backed up with application-aware processing show No. For a Veeam Agent job
    object that is a protection group, cluster or Active Directory container, Yes means
    that one of the machines listed for it in SettingsFrom has one.
    This can be slow on large servers.

.PARAMETER PassThru
    Also writes the result objects to the pipeline.

.EXAMPLE
    .\Get-VeeamGuestProcessingAudit.ps1

.EXAMPLE
    .\Get-VeeamGuestProcessingAudit.ps1 -Server vbr01.corp.local -Credential (Get-Credential) -DetectSql

.EXAMPLE
    .\Get-VeeamGuestProcessingAudit.ps1 -JobName 'SQL*' -IncludeDisabled -PassThru | Out-GridView

.NOTES
    Requires PowerShell 7 (Veeam 13.1 needs PowerShell 7.6 or later) and the
    Veeam Backup & Replication v13 console (Veeam.Backup.PowerShell module) on the
    machine that runs the script. The console build must match the backup server build.

    InLastSession and LastResult come from the latest job run: the latest full session
    and the retry, Quick Backup or planned failover sessions after it. A VM that is in a container but
    shows InLastSession False was added after that run, or it did not run.
#>
[CmdletBinding()]
param(
    [string]$Server = 'localhost',
    [pscredential]$Credential,
    [ValidateRange(1, 65535)]
    [int]$Port,
    [switch]$ForceAcceptTlsCertificate,
    [string[]]$JobName = @('*'),
    [string]$OutputPath,
    [switch]$IncludeDisabled,
    [switch]$DetectSql,
    [switch]$PassThru
)

$ErrorActionPreference = 'Stop'

#region Helpers

# Reads a dotted property path, for example 'VssSnapshotOptions.IsCopyOnly'.
# Returns $null when a step is missing, so one script works across Veeam builds.
function Get-SafeValue {
    param($InputObject, [string]$Path)

    $current = $InputObject
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $current) { return $null }
        try {
            $property = $current.PSObject.Properties[$part]
            if ($null -eq $property) { return $null }
            $current = $property.Value
        }
        catch { return $null }
    }
    return $current
}

# Include, Exclude, VssChild ("Applications..." entry) or GuestFSIndexingChild ("Indexing..." entry).
function Get-JobObjectRole {
    param($ObjectInJob)

    $role = Get-SafeValue $ObjectInJob 'Type'
    if (-not $role) { $role = Get-SafeValue $ObjectInJob 'Info.Type' }
    return [string]$role
}

function Get-PlatformName {
    param([string]$Value)

    if ($Value -match 'HyperV|^E?Hv') { return 'HyperV' }
    return 'VMware'
}

# Key of a VM or container in the inventory and in the Veeam database: '<host id>_<reference>'.
function Get-ObjectKey {
    param($HostId, $Reference)

    $hostText = [string]$HostId
    $referenceText = [string]$Reference
    if (-not $hostText -or -not $referenceText -or $hostText -eq [string][guid]::Empty) { return '' }
    return ('{0}_{1}' -f $hostText, $referenceText).ToLowerInvariant()
}

# Key of a Veeam database object (CHierarchyObj).
function Get-HierarchyObjectKey {
    param($Object)

    $reference = Get-SafeValue $Object 'ObjectId'
    if (-not $reference) { $reference = Get-SafeValue $Object 'Info.ObjectId' }
    return Get-ObjectKey (Get-SafeValue $Object 'HostId') $reference
}

# Reads a job object once. Each read of .Object or .Name asks the backup server again
# (about 15 ms), so the script keeps the values it needs.
function Get-JobObjectInfo {
    param($ObjectInJob, [string]$JobPlatform)

    $object = Get-SafeValue $ObjectInJob 'Object'
    $name = [string](Get-SafeValue $object 'Name')
    if (-not $name) { $name = [string](Get-SafeValue $ObjectInJob 'Name') }
    $viType = [string](Get-SafeValue $object 'ViType')
    $objectType = [string](Get-SafeValue $object 'Type')

    # Object.Type is 'VM' also for datacenters, clusters, datastores, resource pools and vApps.
    # Object.ViType tells them apart ('VirtualMachine', or 'Virtual machine' for Hyper-V).
    # vCenter, ESXi and Hyper-V host objects have no ViType.
    if ($viType) { $isVm = $viType -match '^(VirtualMachine|Virtual ?machine)$' }
    else {
        $isVm = $objectType -match '(vm|virtual ?machine|^Template)$' -or
                [string](Get-SafeValue $ObjectInJob 'TypeDisplayName') -match '(vm|virtual ?machine)$'
    }

    $location = ([string](Get-SafeValue $ObjectInJob 'Location')).TrimEnd('\')
    $locationPrefix = ''
    if ($location -and ($location -ieq $name -or $location.EndsWith("\$name", [StringComparison]::OrdinalIgnoreCase))) {
        $locationPrefix = $location
    }
    elseif ($location) { $locationPrefix = "$location\$name" }

    $platformValue = [string](Get-SafeValue $object 'Platform')
    $platform = if ($platformValue) { Get-PlatformName $platformValue } else { $JobPlatform }

    return [pscustomobject]@{
        Source         = $ObjectInJob
        Name           = $name
        Role           = Get-JobObjectRole $ObjectInJob
        IsVm           = $isVm
        IsTemplate     = $objectType -eq 'Template'
        IsUnix         = (Get-SafeValue $object 'GuestInfo.IsUnixBased') -eq $true
        IsTag          = $viType -match '^(Tag|Category|MultiTag)$'
        ViType         = $viType
        ObjectType     = $objectType
        ObjectRef      = [string](Get-SafeValue $object 'ObjectId')
        Key            = Get-HierarchyObjectKey $object
        LocationPrefix = $locationPrefix
        Platform       = $platform
        Options        = Get-ObjectVssOptions $ObjectInJob
        Settings       = $null
        Prefixes       = @()
        VmKeys         = [System.Collections.Generic.HashSet[string]]::new()
        CanExpand      = $false
        MayMissVms     = $false
        MissReason     = ''
        IsPartial      = $false
        Rank           = 0
    }
}

function Get-JobVssOptions {
    param($Job)

    try {
        $options = $Job.GetVssOptions()
        if ($options) { return $options }
    }
    catch { }
    $options = Get-SafeValue $Job 'VssOptions'
    if ($options) { return $options }
    try { return Get-VBRJobVSSOptions -Job $Job -ErrorAction Stop -WarningAction SilentlyContinue }
    catch { return $null }
}

function Get-ObjectVssOptions {
    param($ObjectInJob)

    try {
        $options = Get-VBRJobObjectVssOptions -ObjectInJob $ObjectInJob -ErrorAction Stop -WarningAction SilentlyContinue
        if ($options) { return $options }
    }
    catch { }
    return Get-SafeValue $ObjectInJob 'VssOptions'
}

# Turns Veeam VSS options into readable guest processing and SQL settings.
function ConvertTo-GuestSettings {
    param($VssOptions, [bool]$JobAapEnabled)

    $settings = [ordered]@{
        GuestProcessing      = 'On'
        ProcessingMode       = ''
        TransactionLogs      = ''
        SqlLogHandling       = ''
        SqlLogBackupEveryMin = ''
        SqlLogRetention      = ''
    }

    if (-not $JobAapEnabled) {
        $settings.GuestProcessing = 'Off (job)'
        $settings.SqlLogHandling = 'N/A'
    }
    elseif ($null -eq $VssOptions) {
        $settings.GuestProcessing = 'Unknown'
        $settings.SqlLogHandling = 'Unknown'
    }
    elseif ((Get-SafeValue $VssOptions 'Enabled') -eq $false) {
        $settings.GuestProcessing = 'Off (VM setting)'
        $settings.ProcessingMode = 'Disable application processing'
        $settings.SqlLogHandling = 'N/A'
    }
    elseif ((Get-SafeValue $VssOptions 'VssSnapshotOptions.ApplicationProcessingEnabled') -eq $false) {
        # Veeam runs application processing only when Enabled and ApplicationProcessingEnabled are both on.
        $settings.GuestProcessing = 'Off (VM setting)'
        $settings.ProcessingMode = 'Application processing off (ApplicationProcessingEnabled = False)'
        $settings.SqlLogHandling = 'N/A'
    }
    else {
        $ignoreErrors = ((Get-SafeValue $VssOptions 'VssSnapshotOptions.IgnoreErrors') -eq $true) -or
                        ((Get-SafeValue $VssOptions 'IgnoreErrors') -eq $true)
        $settings.ProcessingMode = if ($ignoreErrors) { 'Try, ignore failures' } else { 'Require successful processing' }

        if ((Get-SafeValue $VssOptions 'VssSnapshotOptions.IsCopyOnly') -eq $true) {
            $settings.TransactionLogs = 'Perform copy only'
            $settings.SqlLogHandling = 'Copy only'
        }
        else {
            $settings.TransactionLogs = 'Process transaction logs'
            $logMode = [string](Get-SafeValue $VssOptions 'SqlBackupOptions.TransactionLogsProcessing')
            switch ($logMode) {
                'TruncateOnlyOnSuccessJob' { $settings.SqlLogHandling = 'Truncate logs' }
                'NeverTruncate' { $settings.SqlLogHandling = 'Do not truncate logs' }
                'Backup' {
                    $settings.SqlLogHandling = 'Back up logs (log shipping)'
                    $settings.SqlLogBackupEveryMin = [string](Get-SafeValue $VssOptions 'SqlBackupOptions.BackupLogsFrequencyMin')
                    if ((Get-SafeValue $VssOptions 'SqlBackupOptions.UseDbBackupRetention') -eq $true) {
                        $settings.SqlLogRetention = 'Until image-level backup is deleted'
                    }
                    else {
                        $days = Get-SafeValue $VssOptions 'SqlBackupOptions.RetainDays'
                        if ($days) { $settings.SqlLogRetention = "Keep last $days days" }
                    }
                }
                '' { $settings.SqlLogHandling = 'Unknown' }
                default { $settings.SqlLogHandling = $logMode }
            }
        }
    }

    $settings.Signature = $settings.Values -join '|'
    return [pscustomobject]$settings
}

# Same fields as ConvertTo-GuestSettings, for the application processing options of a Veeam Agent job.
function ConvertTo-AgentGuestSettings {
    param($AppOptions, [bool]$JobAapEnabled)

    $settings = [ordered]@{
        GuestProcessing      = 'On'
        ProcessingMode       = ''
        TransactionLogs      = ''
        SqlLogHandling       = ''
        SqlLogBackupEveryMin = ''
        SqlLogRetention      = ''
    }

    if (-not $JobAapEnabled) {
        $settings.GuestProcessing = 'Off (job)'
        $settings.SqlLogHandling = 'N/A'
    }
    elseif ((Get-SafeValue $AppOptions 'Enabled') -eq $false) {
        $settings.GuestProcessing = 'Off (VM setting)'
        $settings.ProcessingMode = 'Disable application processing'
        $settings.SqlLogHandling = 'N/A'
    }
    else {
        $settings.ProcessingMode = if ((Get-SafeValue $AppOptions 'IgnoreErrors') -eq $true) { 'Try, ignore failures' } else { 'Require successful processing' }
        if ([string](Get-SafeValue $AppOptions 'GeneralTransactionLogAction') -eq 'PerformCopyOnly') {
            $settings.TransactionLogs = 'Perform copy only'
            $settings.SqlLogHandling = 'Copy only'
        }
        else {
            $settings.TransactionLogs = 'Process transaction logs'
            $sql = Get-SafeValue $AppOptions 'SQLProcessingOptions'
            $logMode = [string](Get-SafeValue $sql 'TransactionAction')
            switch ($logMode) {
                'Truncate' { $settings.SqlLogHandling = 'Truncate logs' }
                'NotTruncate' { $settings.SqlLogHandling = 'Do not truncate logs' }
                'BackupPeriodically' {
                    $settings.SqlLogHandling = 'Back up logs (log shipping)'
                    $settings.SqlLogBackupEveryMin = [string](Get-SafeValue $sql 'LogBackupPeriod')
                    $settings.SqlLogRetention = if ([string](Get-SafeValue $sql 'LogRetainAction') -eq 'KeepOnlyLastDays') {
                        "Keep last $(Get-SafeValue $sql 'LogRetainPeriod') days"
                    }
                    else { 'Until image-level backup is deleted' }
                }
                '' { $settings.SqlLogHandling = 'Unknown' }
                default { $settings.SqlLogHandling = $logMode }
            }
        }
    }

    $settings.Signature = $settings.Values -join '|'
    return [pscustomobject]$settings
}

$script:InventoryCache = @{}

# Reads VMs and containers from the virtual infrastructure, once per platform and view.
# Each path starts with its view, for example 'VMsAndTemplates|vcenter01\Datacenter1\Folder1\vm1',
# so a folder and a cluster with the same name do not mix.
function Get-VmInventory {
    param([string]$Platform, [string[]]$Views)

    $inventory = $script:InventoryCache[$Platform]
    if (-not $inventory) {
        $inventory = [pscustomobject]@{
            Entities            = [System.Collections.Generic.List[object]]::new()
            ByKey               = @{}
            ByName              = @{}
            Containers          = @{}
            TaggedContainerTags = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            ClusterPaths        = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            TagParents          = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            ViewProblems        = @{}
            DownHosts           = @{}
            LoadedViews         = [System.Collections.Generic.List[string]]::new()
            PathList            = @()
            WarnedEmpty         = $false
        }
        $script:InventoryCache[$Platform] = $inventory
    }

    $newViews = @($Views | Where-Object { -not $inventory.LoadedViews.Contains($_) })
    if (-not $newViews.Count) { return $inventory }

    foreach ($view in $newViews) {
        $inventory.LoadedViews.Add($view)
        Write-Host "  Reading $Platform inventory ($view view)..."
        $viewSwitch = @{ $view = $true }
        $viewWarnings = $null
        $problems = [System.Collections.Generic.List[string]]::new()
        $items = @()
        try {
            if ($Platform -eq 'HyperV') {
                $items = @(Find-VBRHvEntity @viewSwitch -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable viewWarnings)
            }
            else {
                $items = @(Find-VBRViEntity @viewSwitch -ErrorAction Stop -WarningAction SilentlyContinue -WarningVariable viewWarnings)
            }
        }
        catch { $problems.Add($_.Exception.Message) }

        # Veeam returns what it could read and only warns about a vCenter, host or cluster it could not read.
        foreach ($warning in @($viewWarnings)) {
            $message = [string](Get-SafeValue $warning 'Message')
            if (-not $message) { $message = [string]$warning }
            if ($message -and -not $problems.Contains($message)) { $problems.Add($message) }
        }

        $topHosts = @{}
        $hostsWithItems = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($item in $items) {
            $path = [string](Get-SafeValue $item 'Path')
            if (-not $path) { continue }
            $name = [string](Get-SafeValue $item 'Name')
            $hostId = ([string](Get-SafeValue $item 'ConnHostId')).ToLowerInvariant()
            if ($hostId -and $path.Contains('\')) { $null = $hostsWithItems.Add($hostId) }
            elseif ($hostId) { $topHosts[$hostId] = $name }
            $key = Get-ObjectKey (Get-SafeValue $item 'ConnHostId') (Get-SafeValue $item 'Reference')
            $typeName = $item.GetType().Name
            # A host in a cluster is listed as '<cluster>\<host>' without the cluster's VMs: they are under the cluster.
            if ($typeName -match '^C(Hv|Vi)ClusterItem$') { $null = $inventory.ClusterPaths.Add($path) }
            if ($typeName -match 'VmItem$') {
                if (-not $key) { $key = 'name:' + $name.ToLowerInvariant() }
                $entity = $inventory.ByKey[$key]
                if (-not $entity) {
                    $entity = [pscustomobject]@{
                        Key        = $key
                        Name       = $name
                        IsTemplate = (Get-SafeValue $item 'IsTemplate') -eq $true
                        IsUnix     = (Get-SafeValue $item 'GuestInfo.IsUnixBased') -eq $true
                        Paths      = [System.Collections.Generic.List[string]]::new()
                    }
                    $inventory.ByKey[$key] = $entity
                    $inventory.Entities.Add($entity)
                }
                if (-not $entity.Paths.Contains("$view|$path")) { $entity.Paths.Add("$view|$path") }
            }
            elseif ($key) {
                # A tag can be set on a folder or a cluster. The Tags view then lists that container, not its VMs.
                # That path under the tag is not where the container is, so it is not kept as a path of the container.
                if ($view -eq 'Tags' -and $typeName -notmatch 'Tag|Categor') {
                    if ($path.Contains('\')) { $null = $inventory.TaggedContainerTags.Add("Tags|$($path.Substring(0, $path.LastIndexOf('\')))") }
                    continue
                }
                # The category of each tag that the Tags view lists. A category with no tag listed can still hold VMs (see Resolve-ContainerVms).
                if ($view -eq 'Tags' -and $typeName -match 'Tag' -and $path.Contains('\')) { $null = $inventory.TagParents.Add($path.Substring(0, $path.LastIndexOf('\'))) }
                if (-not $inventory.Containers.ContainsKey($key)) { $inventory.Containers[$key] = [System.Collections.Generic.List[string]]::new() }
                if (-not $inventory.Containers[$key].Contains("$view|$path")) { $inventory.Containers[$key].Add("$view|$path") }
            }
        }

        if ($problems.Count) {
            # A vCenter, host or cluster that could not be read is listed with nothing under it.
            # In the other views a server can be empty and still online (a vCenter without tags, a Hyper-V host
            # without VM groups), so there a server counts as down only when the main view found it down too.
            $mainView = if ($Platform -eq 'HyperV') { 'HostsAndVMs' } else { 'HostsAndClusters' }
            $mainDown = $inventory.DownHosts[$mainView]
            $downHosts = @($topHosts.Keys | Where-Object {
                    -not $hostsWithItems.Contains($_) -and ($view -eq $mainView -or ($mainDown -and $mainDown.Contains($_)))
                })
            $where = if ($downHosts.Count) {
                $inventory.DownHosts[$view] = [System.Collections.Generic.HashSet[string]]::new([string[]]$downHosts)
                "server(s) $(($downHosts | ForEach-Object { $topHosts[$_] }) -join ', ')"
            }
            else {
                $inventory.ViewProblems[$view] = @($problems)
                'a server'
            }
            Write-Warning "Part of the $Platform inventory ($view view) could not be read, on $($where): $($problems -join ' | ') Is every vCenter, host and cluster online, with valid credentials? VMs on that server are matched to their container by a best guess (rows marked '(assumed)' or Unknown). Fix the connection and run the script again."
        }
    }

    # One line per VM path, '<view>|<path><tab><index>', so a container is expanded with one -like pass.
    $pathList = [System.Collections.Generic.List[string]]::new()
    $inventory.ByName = @{}
    for ($index = 0; $index -lt $inventory.Entities.Count; $index++) {
        $entity = $inventory.Entities[$index]
        foreach ($viewPath in $entity.Paths) { $pathList.Add("$viewPath`t$index") }
        $nameKey = $entity.Name.ToLowerInvariant()
        if (-not $inventory.ByName.ContainsKey($nameKey)) { $inventory.ByName[$nameKey] = [System.Collections.Generic.List[object]]::new() }
        $inventory.ByName[$nameKey].Add($entity)
    }
    $inventory.PathList = $pathList.ToArray()

    if ($inventory.Entities.Count -eq 0 -and -not $inventory.WarnedEmpty) {
        $inventory.WarnedEmpty = $true
        Write-Warning "The $Platform inventory has no VMs (is the vCenter or host online?). VMs in containers are matched to their container by a best guess. See the SettingsFrom column."
    }
    return $inventory
}

# The inventory views that can show the VMs of a container.
function Get-ContainerViews {
    param($Info)

    if ($Info.Platform -eq 'HyperV') {
        if ($Info.ViType -match 'Tag') { return @('HostsAndVMs', 'Tags') }
        if ($Info.ViType -match 'group' -and $Info.ViType -notmatch 'Host') { return @('HostsAndVMs', 'VMGroups') }
        return @('HostsAndVMs')
    }
    if ($Info.IsTag) { return @('HostsAndClusters', 'VMsAndTemplates', 'Tags') }
    if ($Info.ViType -match '^(Datastore|StoragePod)$') { return @('HostsAndClusters', 'VMsAndTemplates', 'DatastoresAndVMs') }
    return @('HostsAndClusters', 'VMsAndTemplates')
}

# The host ID part of a key ('<host id>_<reference>').
function Get-KeyHost {
    param([string]$Key)

    if (-not $Key -or $Key.StartsWith('name:') -or -not $Key.Contains('_')) { return '' }
    return $Key.Substring(0, $Key.IndexOf('_'))
}

# True when the inventory view was read fully, or at least for this host (vCenter, ESXi or Hyper-V host or cluster).
function Test-InventoryViewComplete {
    param($Inventory, [string]$View, [string]$HostId)

    if ($Inventory.Entities.Count -eq 0) { return $false }
    $views = if ($View -eq '*') { @($Inventory.LoadedViews) } else { @($View) }
    foreach ($name in $views) {
        if (-not $Inventory.LoadedViews.Contains($name) -or $Inventory.ViewProblems.ContainsKey($name)) { return $false }
        $downHosts = $Inventory.DownHosts[$name]
        if ($downHosts -and $downHosts.Count -and (-not $HostId -or $downHosts.Contains($HostId))) { return $false }
    }
    return $true
}

# Finds the inventory paths of a container and the VMs under them.
function Resolve-ContainerVms {
    param($Info, $Inventory)

    if ($Info.Key -and $Inventory.Containers.ContainsKey($Info.Key)) {
        # The live inventory path. The stored Location is not updated when the container
        # is renamed or moved in vCenter.
        $Info.Prefixes = @($Inventory.Containers[$Info.Key])
    }
    elseif ($Info.LocationPrefix) {
        $Info.Prefixes = @("*|$($Info.LocationPrefix)")
    }

    $complete = $Info.Prefixes.Count -gt 0
    foreach ($prefix in $Info.Prefixes) {
        $view, $path = $prefix.Split('|', 2)
        if (-not (Test-InventoryViewComplete -Inventory $Inventory -View $view -HostId (Get-KeyHost $Info.Key))) { $complete = $false }
        $pattern = '{0}|{1}\*' -f $view, [WildcardPattern]::Escape($path)
        foreach ($hit in @($Inventory.PathList -like $pattern)) {
            $entity = $Inventory.Entities[[int]$hit.Substring($hit.LastIndexOf("`t") + 1)]
            $null = $Info.VmKeys.Add($entity.Key)
        }
    }

    $Info.CanExpand = $complete -and $Info.ViType -ne 'MultiTag'
    # Why the inventory can miss VMs of this container. An incomplete inventory is reported once, by Get-VmInventory.
    $Info.MissReason = ''
    if ($Info.ViType -eq 'MultiTag') { $Info.MissReason = 'it is a tags combination' }
    elseif (-not $Info.Prefixes.Count) { $Info.MissReason = 'it has no inventory path' }
    elseif (@($Info.Prefixes | Where-Object {
                $prefix = $_
                # The tag itself, or the category of a tag, that is set on a folder or cluster.
                @($Inventory.TaggedContainerTags | Where-Object { $_ -eq $prefix -or $_.StartsWith("$prefix\", [StringComparison]::OrdinalIgnoreCase) }).Count
            }).Count) {
        # The Tags view lists a tagged folder or cluster, not its VMs.
        $Info.MissReason = 'the tag is also set on a folder or cluster'
    }
    elseif (($Info.ViType -in 'Host', 'HostSystem' -or $Info.ObjectType -in 'HvServer', 'ESXi') -and @($Info.Prefixes | Where-Object {
                $path = $_.Split('|', 2)[1]
                $path.Contains('\') -and $Inventory.ClusterPaths.Contains($path.Substring(0, $path.LastIndexOf('\')))
            }).Count) {
        # The inventory lists the VMs of a cluster under the cluster, not under the host that runs them
        # (Hyper-V cluster node: ViType 'Host'; ESXi host in a vSphere cluster: ViType 'HostSystem').
        $Info.MissReason = 'it is a host in a cluster'
    }
    elseif ($complete -and $Info.ViType -eq 'Category' -and -not @($Info.Prefixes | Where-Object {
                $categoryPath = $_.Split('|', 2)[1]
                @($Inventory.TagParents | Where-Object { $_ -eq $categoryPath -or $_.StartsWith("$categoryPath\", [StringComparison]::OrdinalIgnoreCase) }).Count
            }).Count) {
        # Veeam lists a tag category even when it could not read its tags: CViTagsHierarchy.SafeFillTagsByCategory
        # catches the QueryTagsByCategory error and only logs it, with no warning. So a category with no tag
        # in the inventory can still hold VMs. Only on a fully read server: a server or view that the inventory
        # could not read is reported once, by Get-VmInventory.
        $Info.MissReason = 'the inventory lists no tags in it'
    }
    $Info.MayMissVms = (-not $Info.CanExpand) -or [bool]$Info.MissReason
    # A container that the inventory found by its key and read fully, with no VMs, is empty: it holds no VMs.
    $foundByKey = $Info.Key -and $Inventory.Containers.ContainsKey($Info.Key)
    $Info.IsPartial = $Info.MayMissVms -or ($Info.VmKeys.Count -eq 0 -and -not $foundByKey)
}

# How deep Veeam ranks a job object. Veeam uses the settings of the deepest object that holds the VM.
function Get-ObjectRank {
    param($Info)

    if ($Info.IsVm) { return 1000 }
    if ($Info.Platform -eq 'HyperV') {
        switch -Regex ("$($Info.ViType)|$($Info.ObjectType)") {
            'Tag' { return 900 }
            '^VM group' { return 800 }
            'HvServer|^Host\|' { return 700 }
            'HvCluster|^Cluster\|' { return 600 }
            'Scvmm|SCVMM' { return 400 }
            'volume|Csv|CSV' { return 300 }
        }
        return 500
    }
    $depth = 0
    foreach ($prefix in $Info.Prefixes) { $depth = [Math]::Max($depth, $prefix.Split('|', 2)[1].Split('\').Count) }
    # Veeam: a tags combination is one deeper than its deepest tag, a tag one deeper than its category.
    if ($Info.ViType -eq 'MultiTag') { return 790 }
    if ($Info.IsTag) { return 750 + $depth }
    # The VMs and Templates view also lists datacenter folders (group-d...), so the reference decides when there is one.
    $isVmFolder = $Info.ViType -eq 'Folder' -and ($Info.ObjectRef -like 'group-v*' -or
        (-not $Info.ObjectRef -and @($Info.Prefixes | Where-Object { $_ -like 'VMsAndTemplates|*' }).Count -gt 0))
    if ($isVmFolder) { return 600 + $depth }
    return 500 + $depth
}

# Set per job: $true for Hyper-V jobs. A Hyper-V VM keeps its ID under a host and under its cluster
# (two host IDs). A VMware VM reference (vm-123) is unique only within one vCenter.
$script:SameVmOnOtherHost = $false

# True when one of the keys has the same reference (the part after the host ID) as the key.
function Test-SameReference {
    param($Keys, [string]$Key)

    # On VMware the same reference on another vCenter is another VM.
    if (-not $script:SameVmOnOtherHost) { return $false }
    $reference = $Key.Substring($Key.IndexOf('_') + 1)
    foreach ($candidate in $Keys) {
        if ($candidate.Substring($candidate.IndexOf('_') + 1) -eq $reference) { return $true }
    }
    return $false
}

# True when the job object is this VM, or a container that holds it.
function Test-ObjectHoldsVm {
    param($Info, $Vm)

    if ($Info.IsVm) {
        if ($Info.Key -and $Vm.Keys.Count) {
            if ($Vm.Keys.Contains($Info.Key)) { return $true }
            # Same VM reference and name, but another host ID (for example a Hyper-V host and its cluster).
            return $Info.Name -ieq $Vm.Name -and (Test-SameReference -Keys $Vm.Keys -Key $Info.Key)
        }
        return $Info.Name -ieq $Vm.Name
    }
    foreach ($key in $Vm.Keys) {
        if ($Info.VmKeys.Contains($key)) { return $true }
    }
    return $false
}

# Job objects that are single VMs, by name and by VM reference. With it, each VM is compared only
# with its own entries and the containers, not with every VM of the job.
function New-VmObjectLookup {
    param([object[]]$Infos)

    $lookup = @{}
    foreach ($info in $Infos) {
        if (-not $info.IsVm) { continue }
        $lookupKeys = @('name:' + ([string]$info.Name).ToLowerInvariant())
        if ($info.Key) { $lookupKeys += 'ref:' + $info.Key.Substring($info.Key.IndexOf('_') + 1) }
        foreach ($lookupKey in $lookupKeys) {
            if (-not $lookup.ContainsKey($lookupKey)) { $lookup[$lookupKey] = [System.Collections.Generic.List[object]]::new() }
            if (-not $lookup[$lookupKey].Contains($info)) { $lookup[$lookupKey].Add($info) }
        }
    }
    return $lookup
}

# Containers by the keys of the VMs they hold: VM key -> positions in the list (job order). With it, each VM
# is compared only with the containers that hold it, not with every container of the job.
function New-ContainerLookup {
    param([object[]]$Containers)

    $lookup = [hashtable]::new([StringComparer]::Ordinal)
    for ($position = 0; $position -lt $Containers.Count; $position++) {
        foreach ($key in $Containers[$position].VmKeys) {
            if ($null -eq $lookup[$key]) { $lookup[$key] = @($position) } else { $lookup[$key] += $position }
        }
    }
    return $lookup
}

# The job objects that can hold the VM: the VM entries with its name or reference, then the containers.
function Get-CandidateObjects {
    param($Vm, [hashtable]$Lookup, [object[]]$Containers, [hashtable]$ContainerLookup)

    $found = [System.Collections.Generic.List[object]]::new()
    $lookupKeys = @('name:' + ([string]$Vm.Name).ToLowerInvariant()) + @($Vm.Keys | ForEach-Object { 'ref:' + $_.Substring($_.IndexOf('_') + 1) })
    foreach ($lookupKey in $lookupKeys) {
        if (-not $Lookup.ContainsKey($lookupKey)) { continue }
        foreach ($info in $Lookup[$lookupKey]) { if (-not $found.Contains($info)) { $found.Add($info) } }
    }
    # The containers that hold one of the VM's keys (the test in Test-ObjectHoldsVm), in job order.
    $positions = [int[]]@(foreach ($key in $Vm.Keys) { $hit = $ContainerLookup[$key]; if ($null -ne $hit) { $hit } })
    if ($positions.Count -gt 1) { [Array]::Sort($positions) }
    $previous = -1
    foreach ($position in $positions) {
        if ($position -ne $previous) { $found.Add($Containers[$position]) }
        $previous = $position
    }
    return , $found.ToArray()
}

# Picks the job object whose settings Veeam uses for the VM: the deepest object that holds it.
# At the same depth an "Applications..." entry (VssChild) wins.
function Select-SettingsObject {
    param($Vm, [object[]]$Candidates)

    $matching = @($Candidates | Where-Object { Test-ObjectHoldsVm $_ $Vm })
    if (-not $matching.Count) { return $null }
    $best = $matching | Sort-Object -Stable -Property @{ Expression = 'Rank'; Descending = $true },
        @{ Expression = { if ($_.Role -eq 'VssChild') { 0 } else { 1 } } } | Select-Object -First 1
    # A VM entry always wins over a container. Only flag containers that hold the VM with other settings.
    $others = @(if (-not $best.IsVm) {
            $matching | Where-Object {
                -not [object]::ReferenceEquals($_, $best) -and $_.Settings.Signature -ne $best.Settings.Signature
            }
        })
    return [pscustomobject]@{ Object = $best; Others = $others }
}

function Get-SettingsSourceText {
    param($Info)

    if ($Info.Role -eq 'VssChild' -and $Info.IsVm) { return 'Per-VM (Applications settings)' }
    if ($Info.Role -eq 'VssChild') { return "Applications settings: $($Info.Name)" }
    if ($Info.IsVm) { return 'VM object in job' }
    return "Container: $($Info.Name)"
}

# One record per VM. A VM can have more than one key (from the job object, the inventory
# and the session), and the record is found by any of them.
function Add-CandidateVm {
    param([System.Collections.Generic.List[object]]$List, [hashtable]$Index, [string]$Name, [string]$Key)

    $vm = $null
    if ($Key) { $vm = $Index[$Key] }
    $nameKey = 'name:' + $Name.ToLowerInvariant()
    if (-not $vm) {
        # Without a key on one side, a VM with the same name is the same VM. So is a VM with the same
        # name and reference under another host ID (for example a Hyper-V host and its cluster).
        $sameName = @($Index[$nameKey] | Where-Object {
                $_ -and (-not $Key -or $_.Keys.Count -eq 0 -or (Test-SameReference -Keys $_.Keys -Key $Key))
            })
        if ($sameName.Count -eq 1) { $vm = $sameName[0] }
    }
    if (-not $vm) {
        $vm = [pscustomobject]@{
            Name          = $Name
            Keys          = [System.Collections.Generic.HashSet[string]]::new()
            Path          = ''
            InInventory   = $false
            IsTemplate    = $false
            IsUnix        = $false
            InLastSession = $false
            LastResult    = ''
        }
        $List.Add($vm)
        if (-not $Index.ContainsKey($nameKey)) { $Index[$nameKey] = [System.Collections.Generic.List[object]]::new() }
        $Index[$nameKey].Add($vm)
    }
    if ($Key -and -not $Index.ContainsKey($Key)) {
        $null = $vm.Keys.Add($Key)
        $Index[$Key] = $vm
    }
    return $vm
}

function Get-LastJobSession {
    param($Job)

    # $null from these methods means the job never ran.
    try { return $Job.FindLastSession() }
    catch { }
    # Try once more with the same single-session lookup. Do not read all sessions: that is very slow on large servers.
    try { return [Veeam.Backup.Core.CBackupSession]::FindLastByJob($Job.Id) }
    catch {
        Write-Warning "Job '$($Job.Name)': could not read the last session: $($_.Exception.Message)"
        return $null
    }
}

function Get-SessionTasks {
    param($Session)

    try { return @(Get-VBRTaskSession -Session $Session -ErrorAction Stop -WarningAction SilentlyContinue) }
    catch { }
    try { return @($Session.GetTaskSessions()) }
    catch { return @() }
}

# The tasks of the latest job run. A retry, Quick Backup or planned failover session holds only some VMs, and a run that
# failed before "Building VM list" holds none, so the script reads back to the latest full session
# (at most 20 sessions). Tasks of newer sessions come first.
function Get-RecentSessionTasks {
    param($Session, [string]$JobName)

    $all = [System.Collections.Generic.List[object]]::new()
    $foundFull = $false
    for ($step = 0; $Session -and $step -lt 20; $step++) {
        $sessionTasks = @(Get-SessionTasks -Session $Session)
        foreach ($task in $sessionTasks) { $all.Add($task) }
        # A session for some VMs only: retry, health check retry, Quick Backup, Quick Active Full,
        # quick retry, planned failover.
        $isPartial = (Get-SafeValue $Session 'IsRetryMode') -eq $true -or
                     (Get-SafeValue $Session 'IsRecheckRetry') -eq $true -or
                     (Get-SafeValue $Session 'IsQuickBackup') -eq $true -or
                     (Get-SafeValue $Session 'IsQuickActiveFull') -eq $true -or
                     (Get-SafeValue $Session 'IsPartialActiveFull') -eq $true
        if ($sessionTasks.Count -and -not $isPartial) { $foundFull = $true; break }
        try { $Session = [Veeam.Backup.Core.CBackupSession]::FindPrevSessionByCreationTime($Session.Id) }
        catch { $Session = $null }
    }
    if ($all.Count -and -not $foundFull) {
        Write-Warning "Job '$JobName': no full job session was found. The list of VMs from the sessions can be incomplete."
    }
    return $all.ToArray()
}

$script:TaskObjectCache = @{}

# The Veeam database object of a session task (the VM), found by the task's object ID.
function Get-TaskObject {
    param($Task)

    $objectId = [string](Get-SafeValue $Task 'ObjectId')
    if (-not $objectId) { return $null }
    if ($script:TaskObjectCache.ContainsKey($objectId)) { return $script:TaskObjectCache[$objectId] }
    $object = $null
    try { $object = [Veeam.Backup.Core.CHierarchyObj]::Get([guid]$objectId) }
    catch { }
    $script:TaskObjectCache[$objectId] = $object
    return $object
}

function Test-GlobalExclusion {
    param($Vm, [string]$Platform)

    foreach ($key in $Vm.Keys) {
        if ($script:GlobalExclusions.ContainsKey("key|$key")) { return $true }
    }
    $nameKey = "{0}|{1}" -f $Platform, $Vm.Name.ToLowerInvariant()
    if ($script:GlobalExclusions.ContainsKey("namenokey|$nameKey")) { return $true }
    return $Vm.Keys.Count -eq 0 -and $script:GlobalExclusions.ContainsKey("name|$nameKey")
}

function Test-JobNameMatch {
    param([string]$Name)

    foreach ($pattern in $JobName) {
        # An exact job name wins, so 'SQL [PROD] Daily' is not read as a wildcard pattern.
        if ($Name -eq $pattern) { return $true }
        if ($script:ExactJobNames.ContainsKey($pattern)) { continue }
        if ($Name -like $pattern) { return $true }
    }
    return $false
}

function Get-SqlFound {
    param([string]$Name)

    if ($null -eq $script:SqlVms) { return '' }
    if ($script:SqlVms.ContainsKey($Name.ToLowerInvariant())) { return 'Yes' }
    return 'No'
}

function New-ResultRow {
    param(
        [string]$JobName, [string]$JobType, $JobEnabled, [string]$VmName, [string]$VmPath,
        $Settings, [string]$GuestIndexing, [string]$SqlFound, [string]$SettingsFrom,
        [bool]$InLastSession, [string]$LastResult
    )

    return [pscustomobject][ordered]@{
        JobName              = $JobName
        JobType              = $JobType
        JobEnabled           = $JobEnabled
        VMName               = $VmName
        VmPath               = $VmPath
        GuestProcessing      = $Settings.GuestProcessing
        GuestIndexing        = $GuestIndexing
        ProcessingMode       = $Settings.ProcessingMode
        TransactionLogs      = $Settings.TransactionLogs
        SqlLogHandling       = $Settings.SqlLogHandling
        SqlLogBackupEveryMin = $Settings.SqlLogBackupEveryMin
        SqlLogRetention      = $Settings.SqlLogRetention
        SqlFound             = $SqlFound
        SettingsFrom         = $SettingsFrom
        InLastSession        = $InLastSession
        LastResult           = $LastResult
    }
}

# A row for a job that the script could not check.
function New-ErrorRow {
    param([string]$JobName, [string]$JobType, [string]$Message)

    $errorSettings = [pscustomobject]@{ GuestProcessing = 'Unknown'; ProcessingMode = ''; TransactionLogs = ''; SqlLogHandling = ''; SqlLogBackupEveryMin = ''; SqlLogRetention = '' }
    return New-ResultRow -JobName $JobName -JobType $JobType -JobEnabled '' -VmName '[ERROR - job not checked]' -VmPath '' `
        -Settings $errorSettings -GuestIndexing '' -SqlFound '' -SettingsFrom "Error: $Message" -InLastSession $false -LastResult ''
}

# True when the name is this machine: its name, FQDN, an alias or an IP address of it.
function Test-LocalComputer {
    param([string]$Name)

    if ($Name -in 'localhost', '127.0.0.1', '::1', '.' -or $Name -ieq $env:COMPUTERNAME -or $Name -like "$env:COMPUTERNAME.*") { return $true }
    # GetHostAddresses('') gives the addresses of this machine.
    if (-not $Name) { return $false }
    try {
        $localAddresses = @([System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() |
                ForEach-Object { $_.GetIPProperties().UnicastAddresses } | ForEach-Object { $_.Address })
        foreach ($address in [System.Net.Dns]::GetHostAddresses($Name)) {
            if ([System.Net.IPAddress]::IsLoopback($address) -or $localAddresses -contains $address) { return $true }
        }
    }
    catch { }
    return $false
}

# A Veeam option from the registry of this machine, read the way Veeam reads it: a number is off when it is 0;
# a text is off when it is empty, 0 or false and on when it is 1 or true (spaces ignored). Any other value or
# type, or no value, gives the default.
function Get-VeeamBooleanOption {
    param([string]$Name, [bool]$Default)

    $item = Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Veeam\Veeam Backup and Replication' -Name $Name -ErrorAction SilentlyContinue
    if (-not $item) { return $Default }
    $value = $item.$Name
    if ($value -is [int] -or $value -is [long]) { return $value -ne 0 }
    if ($value -is [string]) {
        $text = $value.Trim()
        if ($text -in '', '0', 'false') { return $false }
        if ($text -in '1', 'true') { return $true }
    }
    return $Default
}

#endregion Helpers

$outputFolder = $null
if ($OutputPath -and (Test-Path -LiteralPath $OutputPath -PathType Container)) {
    # A folder: the CSV gets the default file name in it.
    $outputFolder = $OutputPath
    $OutputPath = $null
}
elseif ($OutputPath) {
    # A path that ends with \ or / names a folder, and that folder does not exist (Export-Csv cannot write to it).
    if ($OutputPath -match '[\\/]$') { throw "The output folder does not exist: $OutputPath" }
    $parentFolder = Split-Path $OutputPath -Parent
    if ($parentFolder -and -not (Test-Path -LiteralPath $parentFolder -PathType Container)) {
        throw "The output folder does not exist: $parentFolder"
    }
}

#region Connect

if (-not (Get-Command Get-VBRJob -ErrorAction SilentlyContinue)) {
    try {
        Import-Module Veeam.Backup.PowerShell -DisableNameChecking -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch {
        throw "Cannot load the Veeam PowerShell module (Veeam.Backup.PowerShell). Run this script in PowerShell 7 (Veeam 13.1 needs 7.6 or later) on a machine with the Veeam Backup & Replication v13 console. The console build must match the backup server build. Error: $($_.Exception.Message)"
    }
}

$serverSession = $null
try { $serverSession = Get-VBRServerSession -ErrorAction Stop -WarningAction SilentlyContinue } catch { }

$connectedServer = [string](Get-SafeValue $serverSession 'Server')
# Do not end a session that the script did not open. Disconnect-VBRServer also stops the file-level
# restore sessions started in this PowerShell window. The session shows the server name as typed
# when it was opened (127.0.0.1 for a local connection), so the same server by another name differs.
if ($serverSession -and (($PSBoundParameters.ContainsKey('Server') -and $connectedServer -ine $Server) -or $Credential)) {
    throw "This PowerShell window is already connected to the Veeam backup server '$connectedServer'. The script does not end a session that it did not open. Run the script without -Server and -Credential to use that session, or run Disconnect-VBRServer first."
}

$connectedByScript = $false
if (-not $serverSession) {
    $connect = @{ Server = $Server }
    if ($PSBoundParameters.ContainsKey('Port')) { $connect.Port = $Port }
    if ($Credential) { $connect.Credential = $Credential }
    if ($ForceAcceptTlsCertificate) {
        if ((Get-Command Connect-VBRServer).Parameters.ContainsKey('ForceAcceptTlsCertificate')) {
            $connect.ForceAcceptTlsCertificate = $true
        }
        else {
            Write-Warning 'This Veeam module does not support -ForceAcceptTlsCertificate. The switch is ignored.'
        }
    }
    Write-Host "Connecting to Veeam backup server '$Server'..."
    Connect-VBRServer @connect
    $connectedByScript = $true
    $connectedServer = [string](Get-SafeValue (Get-VBRServerSession) 'Server')
}
if (-not $connectedServer) { $connectedServer = $Server }
Write-Host "Connected to '$connectedServer'."

#endregion Connect

try {
    # Two registry values on the backup server change which job object Veeam takes the settings from.
    # The script can read them only when it runs on the backup server.
    $tagsLast = -not (Get-VeeamBooleanOption -Name 'TagsPriorityOverContainers' -Default $true)
    $credsFirst = Get-VeeamBooleanOption -Name 'PrioritizeSpecialGuestProcessingCreds' -Default $false
    if (($tagsLast -or $credsFirst) -and (Test-LocalComputer $connectedServer)) {
        if ($tagsLast) {
            Write-Warning 'TagsPriorityOverContainers = 0 on this backup server. Veeam then ranks tags below other containers, but this script uses the default order. Check the rows that show a tag in SettingsFrom in the console.'
        }
        if ($credsFirst) {
            Write-Warning 'PrioritizeSpecialGuestProcessingCreds = 1 on this backup server. In VMware jobs Veeam then takes the settings of a job object that has its own guest OS credentials before a deeper object, but this script uses the default order. Check VMware jobs that set credentials for single objects in the console.'
        }
    }

    $script:SqlVms = $null
    if ($DetectSql) {
        Write-Host 'Reading SQL Server application restore points...'
        try {
            $script:SqlVms = @{}
            foreach ($restorePoint in Get-VBRApplicationRestorePoint -SQL -ErrorAction Stop -WarningAction SilentlyContinue) {
                $script:SqlVms[([string]$restorePoint.Name).ToLowerInvariant()] = $true
            }
        }
        catch {
            Write-Warning "Could not read SQL Server restore points. The SqlFound column stays empty. Error: $($_.Exception.Message)"
            $script:SqlVms = $null
        }
    }

    #region Jobs to check

    $allJobs = @(Get-VBRJob -WarningAction SilentlyContinue)
    $script:ExactJobNames = @{}
    foreach ($candidate in $allJobs) { $script:ExactJobNames[[string]$candidate.Name] = $true }
    foreach ($pattern in $JobName) {
        if ($script:ExactJobNames.ContainsKey($pattern)) { continue }
        try { $null = '' -like $pattern }
        catch { throw "-JobName '$pattern' is not a valid wildcard pattern. To match [ or ] literally, put a backtick (``) before it." }
    }
    $matchingJobs = @($allJobs | Where-Object { Test-JobNameMatch ([string]$_.Name) })

    $jobs = [System.Collections.Generic.List[object]]::new()
    $agentJobIds = [System.Collections.Generic.List[string]]::new()
    $copyJobs = [System.Collections.Generic.List[object]]::new()
    $otherJobs = [System.Collections.Generic.List[object]]::new()
    foreach ($candidate in $matchingJobs) {
        $type = [string]$candidate.JobType
        $platformValue = [string](Get-SafeValue $candidate 'BackupPlatform.Platform')
        $isVmPlatform = -not $platformValue -or $platformValue -match 'Vmware|HyperV|^E?Hv|Vcd'
        if ($isVmPlatform -and ($type -in 'Backup', 'Replica' -or ($type -eq 'Copy' -and (Get-SafeValue $candidate 'IsVmCopy') -eq $true))) {
            $jobs.Add($candidate)
        }
        elseif ($type -like 'EpAgent*') { $agentJobIds.Add([string]$candidate.Id) }
        elseif ($type -match 'BackupSync|BackupCopy|^Copy$') { $copyJobs.Add($candidate) }
        else { $otherJobs.Add($candidate) }
    }
    $jobs = @($jobs | Sort-Object Name)
    Write-Host "Found $($jobs.Count) VM backup, replication or VM copy job(s) and $($agentJobIds.Count) Veeam Agent job(s) to check."
    if ($copyJobs.Count) {
        Write-Host "Not checked (backup copy and file copy jobs have no guest processing): $(($copyJobs | ForEach-Object { $_.Name }) -join '; ')"
    }
    if ($otherJobs.Count) {
        $summary = @($otherJobs | ForEach-Object {
                $type = [string](Get-SafeValue $_ 'TypeToString')
                if (-not $type) { $type = [string]$_.JobType }
                "$($_.Name) [$type]"
            }) -join '; '
        Write-Warning "Not checked (this script does not support these job types): $summary. Check their guest processing in the console."
    }

    # Get-VBRJob does not return CDP policies. This script does not audit them.
    try {
        foreach ($policy in @(Get-VBRCDPPolicy -ErrorAction Stop -WarningAction SilentlyContinue)) {
            if (-not (Test-JobNameMatch ([string]$policy.Name))) { continue }
            if ((Get-SafeValue $policy 'ApplicationProcessingEnabled') -eq $true) {
                Write-Warning "CDP policy '$($policy.Name)' has application-aware processing enabled. CDP policies are not in this report. Check its SQL settings (Truncate logs / Do not truncate logs) in the console."
            }
        }
    }
    catch { Write-Warning "Could not read CDP policies. VMs protected by CDP are not in this report. Error: $($_.Exception.Message)" }

    if ($jobs.Count -eq 0 -and $agentJobIds.Count -eq 0) {
        Write-Warning "No VM backup, replication, VM copy or Veeam Agent job matches -JobName '$($JobName -join "', '")'."
        return
    }

    # Global VM exclusions (Inventory > Global VM exclusions). Veeam skips these VMs in every job.
    $script:GlobalExclusions = @{}
    if ($jobs.Count) {
        try {
            foreach ($exclusion in @(Get-VBRVMExclusion -ErrorAction SilentlyContinue -WarningAction SilentlyContinue)) {
                $exclusionPlatform = Get-PlatformName ([string](Get-SafeValue $exclusion 'Platform'))
                $nameKey = '{0}|{1}' -f $exclusionPlatform, ([string]$exclusion.Name).ToLowerInvariant()
                $exclusionKey = ''
                try { $exclusionKey = Get-HierarchyObjectKey ([Veeam.Backup.Core.CHierarchyObj]::Get([guid]$exclusion.Id)) }
                catch { }
                if ($exclusionKey) { $script:GlobalExclusions["key|$exclusionKey"] = $true }
                else { $script:GlobalExclusions["namenokey|$nameKey"] = $true }
                $script:GlobalExclusions["name|$nameKey"] = $true
            }
        }
        catch { Write-Warning "Could not read the global VM exclusions. Globally excluded VMs can show in the result. Error: $($_.Exception.Message)" }
    }

    #endregion Jobs to check

    $rows = [System.Collections.Generic.List[object]]::new()
    $jobNumber = 0

    foreach ($job in $jobs) {
        $jobNumber++
        Write-Progress -Activity 'Checking Veeam jobs' -Status $job.Name -PercentComplete (100 * $jobNumber / $jobs.Count)
        try {
            $jobVss = Get-JobVssOptions -Job $job
            $jobAapOn = ((Get-SafeValue $jobVss 'Enabled') -eq $true) -and
                        ((Get-SafeValue $jobVss 'VssSnapshotOptions.ApplicationProcessingEnabled') -ne $false)
            if (-not $jobAapOn -and -not $IncludeDisabled) {
                Write-Verbose "Skip '$($job.Name)': application-aware processing is off in the job."
                continue
            }
            Write-Host "Checking job '$($job.Name)'..."

            $jobType = [string](Get-SafeValue $job 'TypeToString')
            if (-not $jobType) { $jobType = [string]$job.JobType }
            $jobEnabled = Get-SafeValue $job 'IsScheduleEnabled'
            $jobPlatform = Get-PlatformName ([string](Get-SafeValue $job 'BackupPlatform.Platform'))
            $script:SameVmOnOtherHost = $jobPlatform -eq 'HyperV'
            if ([string](Get-SafeValue $job 'BackupPlatform.Platform') -match 'Vcd') {
                Write-Warning "Job '$($job.Name)' is a VMware Cloud Director job. The script cannot read the Cloud Director inventory, so its VMs come from the job sessions only and SettingsFrom can be a best guess. Check this job in the console."
            }
            $jobIndexTypes = @('GuestFSIndexingType', 'LinGuestFSIndexingType' | Where-Object {
                    [string](Get-SafeValue $jobVss $_) -notin '', 'None'
                })
            $guestIndexing = if ($jobIndexTypes.Count) { 'On' } else { 'Off' }

            $infos = @(Get-VBRJobObject -Job $job -WarningAction SilentlyContinue | Where-Object {
                    (Get-JobObjectRole $_) -in 'Include', 'Exclude', 'VssChild', 'GuestFSIndexingChild'
                } | ForEach-Object { Get-JobObjectInfo -ObjectInJob $_ -JobPlatform $jobPlatform })
            foreach ($info in $infos) { $info.Settings = ConvertTo-GuestSettings -VssOptions $info.Options -JobAapEnabled $jobAapOn }
            $includes = @($infos | Where-Object Role -eq 'Include')
            $excludes = @($infos | Where-Object Role -eq 'Exclude')
            $vssChildren = @($infos | Where-Object Role -eq 'VssChild')
            $indexChildren = @($infos | Where-Object Role -eq 'GuestFSIndexingChild')
            $includeContainers = @($includes | Where-Object { -not $_.IsVm })
            # Veeam ranks "Applications..." entries and included objects in one list.
            $settingsObjects = @(@($vssChildren) + @($includes))
            $settingsContainers = @($settingsObjects | Where-Object { -not $_.IsVm })
            $settingsLookup = New-VmObjectLookup -Infos $settingsObjects
            $excludeLookup = New-VmObjectLookup -Infos $excludes
            $excludeContainers = @($excludes | Where-Object { -not $_.IsVm })
            $includeLookup = New-VmObjectLookup -Infos $includes
            $indexObjects = @(@($indexChildren) + @($includes))
            $indexLookup = New-VmObjectLookup -Infos $indexObjects
            $indexContainers = @($indexObjects | Where-Object { -not $_.IsVm })

            # Containers: find their VMs in the inventory.
            $inventory = $null
            $views = @()
            $containerInfos = @($infos | Where-Object { -not $_.IsVm })
            if ($containerInfos.Count) {
                $views = @($containerInfos | ForEach-Object { Get-ContainerViews $_ } | Select-Object -Unique)
                $inventory = Get-VmInventory -Platform $jobPlatform -Views $views
                foreach ($info in $containerInfos) { Resolve-ContainerVms -Info $info -Inventory $inventory }
            }
            foreach ($info in $infos) { $info.Rank = Get-ObjectRank $info }
            # Read once per job, not once per VM.
            $excludeContainerLookup = New-ContainerLookup -Containers $excludeContainers
            $settingsContainerLookup = New-ContainerLookup -Containers $settingsContainers
            $includeContainerLookup = New-ContainerLookup -Containers $includeContainers
            $indexContainerLookup = New-ContainerLookup -Containers $indexContainers
            $mayMissSettingsContainers = @($settingsContainers | Where-Object MayMissVms)
            $mayMissExcludeContainers = @($excludeContainers | Where-Object MayMissVms)
            $hasPartialInclude = @($includeContainers | Where-Object IsPartial).Count -gt 0

            $vmList = [System.Collections.Generic.List[object]]::new()
            $vmIndex = @{}

            # 1. VMs added to the job directly.
            foreach ($info in @($includes | Where-Object IsVm)) {
                $vm = Add-CandidateVm -List $vmList -Index $vmIndex -Name $info.Name -Key $info.Key
                if ($info.IsTemplate) { $vm.IsTemplate = $true }
                if ($info.IsUnix) { $vm.IsUnix = $true }
                if (-not $vm.Path) { $vm.Path = $info.LocationPrefix }
            }

            # 2. VMs inside included containers, from the inventory.
            foreach ($info in $includeContainers) {
                foreach ($key in $info.VmKeys) {
                    $entity = $inventory.ByKey[$key]
                    if ($entity -and -not $entity.IsTemplate) { $null = Add-CandidateVm -List $vmList -Index $vmIndex -Name $entity.Name -Key $entity.Key }
                }
            }

            # 3. VMs from the latest job run.
            $lastSession = Get-LastJobSession -Job $job
            $tasks = @(if ($lastSession) { Get-RecentSessionTasks -Session $lastSession -JobName $job.Name })
            foreach ($task in $tasks) {
                $taskName = [string](Get-SafeValue $task 'Name')
                if (-not $taskName) { continue }
                $taskObject = Get-TaskObject $task
                $vm = Add-CandidateVm -List $vmList -Index $vmIndex -Name $taskName -Key (Get-HierarchyObjectKey $taskObject)
                if ($vm.InLastSession) { continue }  # a newer session already gave the result
                $vm.InLastSession = $true
                $vm.LastResult = [string](Get-SafeValue $task 'Status')
                if ([string](Get-SafeValue $taskObject 'Type') -eq 'Template') { $vm.IsTemplate = $true }
                if ((Get-SafeValue $taskObject 'GuestInfo.IsUnixBased') -eq $true) { $vm.IsUnix = $true }
                if (-not $vm.Path) { $vm.Path = [string](Get-SafeValue $taskObject 'Info.Path') }
            }

            # The inventory tells where each VM is now.
            if ($inventory) {
                foreach ($vm in $vmList) {
                    $entity = $null
                    foreach ($key in @($vm.Keys)) { if (-not $entity) { $entity = $inventory.ByKey[$key] } }
                    if (-not $entity) {
                        $sameName = @($inventory.ByName[$vm.Name.ToLowerInvariant()] | Where-Object {
                                $_ -and ($vm.Keys.Count -eq 0 -or (Test-SameReference -Keys $vm.Keys -Key $_.Key))
                            })
                        if ($sameName.Count -eq 1) { $entity = $sameName[0] }
                    }
                    if (-not $entity) { continue }
                    $vm.InInventory = $true
                    if (-not $vmIndex.ContainsKey($entity.Key)) { $null = $vm.Keys.Add($entity.Key); $vmIndex[$entity.Key] = $vm }
                    if ($entity.IsTemplate) { $vm.IsTemplate = $true }
                    if ($entity.IsUnix) { $vm.IsUnix = $true }
                    $displayPath = @($entity.Paths | Where-Object { $_ -like 'VMsAndTemplates|*' }) + @($entity.Paths) | Select-Object -First 1
                    if ($displayPath) { $vm.Path = $displayPath.Split('|', 2)[1] }
                }
            }
            foreach ($exclude in @($excludes | Where-Object MissReason)) {
                Write-Warning "Job '$($job.Name)': the script cannot fully check the exclusion '$($exclude.Name)' ($($exclude.MissReason)). VMs that are not in the last session are marked in SettingsFrom."
            }
            if ($tasks.Count) {
                foreach ($container in @($settingsContainers | Where-Object MissReason)) {
                    Write-Warning "Job '$($job.Name)': the inventory cannot show every VM of '$($container.Name)' ($($container.MissReason)). VMs added to it after the latest job run can be missing, and its VMs from the job sessions are matched by a best guess."
                }
            }

            $movedOut = [System.Collections.Generic.List[string]]::new()
            foreach ($vm in $vmList) {
                # Guest processing never runs on a VM template.
                if ($vm.IsTemplate) { continue }
                if (@((Get-CandidateObjects -Vm $vm -Lookup $excludeLookup -Containers $excludeContainers -ContainerLookup $excludeContainerLookup) | Where-Object { Test-ObjectHoldsVm $_ $vm }).Count -gt 0) { continue }
                if (Test-GlobalExclusion -Vm $vm -Platform $jobPlatform) { continue }

                $settings = $null
                $settingsObject = $null
                $settingsFrom = ''
                $selection = Select-SettingsObject -Vm $vm -Candidates (Get-CandidateObjects -Vm $vm -Lookup $settingsLookup -Containers $settingsContainers -ContainerLookup $settingsContainerLookup)
                # An "Applications..." entry (VssChild) gives settings only. Veeam builds the VM list from the
                # included objects, so a VM that no included object holds is not in the job.
                if ($selection -and $selection.Object.Role -ne 'Include' -and
                    -not @((Get-CandidateObjects -Vm $vm -Lookup $includeLookup -Containers $includeContainers -ContainerLookup $includeContainerLookup) | Where-Object { Test-ObjectHoldsVm $_ $vm }).Count -and
                    -not $hasPartialInclude) {
                    if ($includeContainers.Count) { $movedOut.Add($vm.Name) }
                    continue
                }
                if ($selection) {
                    $settingsObject = $selection.Object
                    $settingsFrom = Get-SettingsSourceText $settingsObject
                    if ($selection.Others.Count) {
                        $settingsFrom += " (also in: $(($selection.Others | ForEach-Object { $_.Name }) -join ', ') with other settings - check in console)"
                    }
                    # A container that the inventory cannot fully read (for example a tag on a folder) can also hold the VM.
                    $unsure = @($mayMissSettingsContainers | Where-Object {
                            $_.Rank -gt $settingsObject.Rank -and
                            $_.Settings.Signature -ne $settingsObject.Settings.Signature -and -not (Test-ObjectHoldsVm $_ $vm)
                        })
                    if ($unsure.Count) {
                        $settingsFrom += " (check: $(($unsure | ForEach-Object { $_.Name }) -join ', ') can also hold this VM)"
                    }
                }
                elseif ($includeContainers.Count -eq 0) {
                    # The job holds only single VMs and this VM is not one of them:
                    # it was removed from the job after the last run.
                    continue
                }
                else {
                    # The VM is not in any container the inventory could read. If the inventory knows where
                    # the VM is (or read its server fully and the VM is not there), only containers the
                    # inventory cannot fully read can hold it.
                    $vmHost = @($vm.Keys | ForEach-Object { Get-KeyHost $_ } | Where-Object { $_ }) | Select-Object -First 1
                    $located = $vm.InInventory -or ($vmHost -and
                        @($views | Where-Object { -not (Test-InventoryViewComplete -Inventory $inventory -View $_ -HostId $vmHost) }).Count -eq 0)
                    $possible = @(if ($located) {
                            $settingsContainers | Where-Object IsPartial
                        }
                        else { $settingsContainers })
                    # Only an included object can add the VM to the job. An "Applications..." entry gives settings only.
                    if (@($possible | Where-Object Role -eq 'Include').Count -eq 0) {
                        $movedOut.Add($vm.Name)
                        continue
                    }
                    $signatures = @($possible | ForEach-Object { $_.Settings.Signature } | Select-Object -Unique)
                    if ($signatures.Count -eq 1) {
                        $settingsObject = $possible[0]
                        $settingsFrom = "Container: $(($possible | ForEach-Object { $_.Name }) -join ', ') (assumed)"
                    }
                    else {
                        $settings = ConvertTo-GuestSettings -VssOptions $null -JobAapEnabled $jobAapOn
                        $possibleLogs = @($possible | ForEach-Object { $_.Settings.SqlLogHandling } | Select-Object -Unique) -join ' | '
                        $settingsFrom = "Not matched to a job object - check in console (possible: $possibleLogs)"
                    }
                }
                if ($settingsObject) { $settings = $settingsObject.Settings }
                if ($settings.GuestProcessing -like 'Off*' -and -not $IncludeDisabled) { continue }

                # SQL Server processing runs on Windows guests only.
                if ($vm.IsUnix -and $settings.GuestProcessing -eq 'On') {
                    $settings = $settings.PSObject.Copy()
                    $settings.TransactionLogs = ''
                    $settings.SqlLogHandling = 'N/A (Linux guest)'
                    $settings.SqlLogBackupEveryMin = ''
                    $settings.SqlLogRetention = ''
                }

                # Veeam never creates a task for an excluded VM, so a VM in the last session is not excluded.
                if (-not $vm.InLastSession) {
                    $uncheckedExcludes = @(if ($vm.InInventory) { $mayMissExcludeContainers } else { $excludeContainers })
                    if ($uncheckedExcludes.Count) {
                        $settingsFrom += "; not in last session - can be excluded by: $(($uncheckedExcludes | ForEach-Object { $_.Name }) -join ', ')"
                    }
                }

                # Indexing has its own entries ("Indexing..."). The job setting turns it on or off.
                $vmIndexing = $guestIndexing
                if ($guestIndexing -eq 'On') {
                    $indexSelection = Select-SettingsObject -Vm $vm -Candidates (Get-CandidateObjects -Vm $vm -Lookup $indexLookup -Containers $indexContainers -ContainerLookup $indexContainerLookup)
                    if ($indexSelection) {
                        $vmIndexValues = @($jobIndexTypes | ForEach-Object { [string](Get-SafeValue $indexSelection.Object.Options $_) } | Where-Object { $_ })
                        if ($vmIndexValues.Count -and @($vmIndexValues | Where-Object { $_ -ne 'None' }).Count -eq 0) { $vmIndexing = 'Off (VM setting)' }
                    }
                }

                $rows.Add((New-ResultRow -JobName $job.Name -JobType $jobType -JobEnabled $jobEnabled -VmName $vm.Name -VmPath $vm.Path `
                            -Settings $settings -GuestIndexing $vmIndexing -SqlFound (Get-SqlFound $vm.Name) -SettingsFrom $settingsFrom `
                            -InLastSession $vm.InLastSession -LastResult $vm.LastResult))
            }
            if ($movedOut.Count) {
                Write-Warning "Job '$($job.Name)': $($movedOut.Count) VM(s) from the latest job run are not in the job containers now (moved or deleted?), so they are not listed: $($movedOut -join ', ')"
            }

            # Without a session that lists VMs, the VMs of a container that the inventory cannot read are unknown.
            if ($tasks.Count -eq 0) {
                foreach ($container in @($includeContainers | Where-Object IsPartial)) {
                    Write-Warning "Job '$($job.Name)': cannot list every VM in container '$($container.Name)'. No job session lists VMs yet."
                    if ($container.Settings.GuestProcessing -like 'Off*' -and -not $IncludeDisabled) { continue }
                    $rows.Add((New-ResultRow -JobName $job.Name -JobType $jobType -JobEnabled $jobEnabled -VmName "[not expanded] $($container.Name)" `
                                -VmPath $container.LocationPrefix -Settings $container.Settings -GuestIndexing $guestIndexing -SqlFound '' `
                                -SettingsFrom "Container: $($container.Name)" -InLastSession $false -LastResult ''))
                }
            }
        }
        catch {
            Write-Warning "Job '$($job.Name)' was not checked: $($_.Exception.Message)"
            $rows.Add((New-ErrorRow -JobName $job.Name -JobType ([string]$job.JobType) -Message $_.Exception.Message))
        }
    }
    Write-Progress -Activity 'Checking Veeam jobs' -Completed

    #region Veeam Agent jobs

    if ($agentJobIds.Count) {
        $agentJobs = @()
        try {
            $agentJobs = @(Get-VBRComputerBackupJob -ErrorAction Stop -WarningAction SilentlyContinue | Where-Object { [string]$_.Id -in $agentJobIds })
        }
        catch {
            # The cmdlet stops at the first agent job that it cannot read. Read the jobs one by one.
            $agentJobs = @(foreach ($agentJobId in $agentJobIds) {
                    try { Get-VBRComputerBackupJob -Id ([guid]$agentJobId) -ErrorAction Stop -WarningAction SilentlyContinue }
                    catch {
                        $message = $_.Exception.Message
                        $legacyJob = $allJobs | Where-Object { [string]$_.Id -eq $agentJobId } | Select-Object -First 1
                        Write-Warning "Agent job '$($legacyJob.Name)' was not checked: $message"
                        $rows.Add((New-ErrorRow -JobName $legacyJob.Name -JobType 'Agent' -Message $message))
                    }
                })
        }
        foreach ($agentJob in $agentJobs) {
            try {
                # For agent jobs the job VSS flag does not show application-aware processing. This property does.
                $aapOn = (Get-SafeValue $agentJob 'ApplicationProcessingEnabled') -eq $true
                # A Workstation job has no application-aware processing, whatever this property says.
                if ([string](Get-SafeValue $agentJob 'Type') -eq 'Workstation' -and $aapOn) {
                    # Only PowerShell or old data can set it. Say so, so this is not silent.
                    Write-Warning "Agent job '$($agentJob.Name)' is a Workstation job, but its settings say application-aware processing is on. Workstation jobs have no application-aware processing, so the script treats it as off: it lists the job only with -IncludeDisabled, as 'Off (job)'. Check this job in the console."
                    $aapOn = $false
                }
                if (-not $aapOn -and -not $IncludeDisabled) {
                    Write-Verbose "Skip agent job '$($agentJob.Name)': application-aware processing is off in the job."
                    continue
                }
                Write-Host "Checking agent job '$($agentJob.Name)'..."
                $agentJobType = "Agent ($(Get-SafeValue $agentJob 'OSPlatform') $(Get-SafeValue $agentJob 'Type'))"
                $agentIndexing = if ((Get-SafeValue $agentJob 'IndexingEnabled') -eq $true) { 'On' } else { 'Off' }
                $tasks = @()
                $legacyJob = $allJobs | Where-Object { [string]$_.Id -eq [string]$agentJob.Id } | Select-Object -First 1
                if ($legacyJob) {
                    $session = Get-LastJobSession -Job $legacyJob
                    if ($session) { $tasks = @(Get-RecentSessionTasks -Session $session -JobName $agentJob.Name) }
                }
                # Tasks of newer sessions come first. Keep the newest status of each machine.
                $seenMachines = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                $tasks = @($tasks | Where-Object { $seenMachines.Add([string]$_.Name) })
                $allOptions = @(Get-SafeValue $agentJob 'ApplicationProcessingOptions' | Where-Object { $_ })
                # Names are compared like -eq and -contains: case-insensitive, invariant culture.
                $entryNames = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::InvariantCultureIgnoreCase)
                foreach ($entry in $allOptions) { $null = $entryNames.Add([string](Get-SafeValue $entry 'BackupObject.Name')) }
                # Once per job, not once per entry: the tasks of machines without an entry of their own,
                # and the tasks of each machine name (task positions, newest first).
                $unownedTasks = [System.Collections.Generic.List[int]]::new()
                $tasksByName = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[int]]]::new([StringComparer]::InvariantCultureIgnoreCase)
                for ($taskIndex = 0; $taskIndex -lt $tasks.Count; $taskIndex++) {
                    $taskName = [string]$tasks[$taskIndex].Name
                    if (-not $entryNames.Contains($taskName)) { $unownedTasks.Add($taskIndex) }
                    if (-not $tasksByName.ContainsKey($taskName)) { $tasksByName[$taskName] = [System.Collections.Generic.List[int]]::new() }
                    $tasksByName[$taskName].Add($taskIndex)
                }
                foreach ($appOptions in $allOptions) {
                    $objectName = [string](Get-SafeValue $appOptions 'BackupObject.Name')
                    # A computer with its own "Applications..." entry uses that entry, not the one of its protection group.
                    # This entry lists its own tasks and the machines without an entry, in task order.
                    $entryTasks = [System.Collections.Generic.List[int]]::new($unownedTasks)
                    $task = $null
                    if ($tasksByName.ContainsKey($objectName)) {
                        $task = $tasks[$tasksByName[$objectName][0]]
                        foreach ($taskIndex in $tasksByName[$objectName]) {
                            $position = $entryTasks.BinarySearch($taskIndex)
                            if ($position -lt 0) { $entryTasks.Insert(-bnot $position, $taskIndex) }
                        }
                    }
                    $machines = @(foreach ($taskIndex in $entryTasks) { "$($tasks[$taskIndex].Name) ($($tasks[$taskIndex].Status))" }) -join ', '
                    $settings = ConvertTo-AgentGuestSettings -AppOptions $appOptions -JobAapEnabled $aapOn
                    if ($settings.GuestProcessing -like 'Off*' -and -not $IncludeDisabled) { continue }
                    # SQL Server processing runs on Windows machines only. Veeam has no SQL options for other agents.
                    $entryPlatform = [string](Get-SafeValue $appOptions 'OSPlatform')
                    if ($settings.GuestProcessing -eq 'On' -and $entryPlatform -notin '', 'Windows') {
                        $settings.TransactionLogs = ''
                        $settings.SqlLogHandling = "N/A ($entryPlatform guest)"
                    }
                    # One row per agent job object (computer, cluster or protection group). Its machines are in SettingsFrom.
                    $settingsFrom = if ($machines) { "Agent job object. Last session: $machines" } else { 'Agent job object' }
                    # Restore points carry machine names. A protection group, cluster or AD container entry gets Yes when
                    # one of the machines listed for it has a SQL Server restore point. A computer keeps its own result,
                    # also when no session lists it yet (VBRDiscoveredComputer or VBRIndividualComputer).
                    $sqlFound = Get-SqlFound $objectName
                    $backupObject = Get-SafeValue $appOptions 'BackupObject'
                    $isComputer = $backupObject -and ($backupObject.PSObject.TypeNames -match '\.VBR(Discovered|Individual)Computer$')
                    if ($sqlFound -eq 'No' -and -not $isComputer -and -not $tasksByName.ContainsKey($objectName)) {
                        foreach ($taskIndex in $entryTasks) {
                            if ((Get-SqlFound ([string]$tasks[$taskIndex].Name)) -eq 'Yes') { $sqlFound = 'Yes'; break }
                        }
                    }
                    $rows.Add((New-ResultRow -JobName $agentJob.Name -JobType $agentJobType -JobEnabled (Get-SafeValue $agentJob 'JobEnabled') `
                                -VmName $objectName -VmPath '' -Settings $settings -GuestIndexing $agentIndexing -SqlFound $sqlFound `
                                -SettingsFrom $settingsFrom -InLastSession ([bool]$task) -LastResult $(if ($task) { [string]$task.Status } else { '' })))
                }
            }
            catch {
                Write-Warning "Agent job '$($agentJob.Name)' was not checked: $($_.Exception.Message)"
                $rows.Add((New-ErrorRow -JobName $agentJob.Name -JobType 'Agent' -Message $_.Exception.Message))
            }
        }
    }

    #endregion Veeam Agent jobs

    $result = @($rows | Sort-Object JobName, VMName)
    if ($result.Count -eq 0) {
        Write-Host 'No VMs with guest processing were found.'
        return
    }

    $result | Format-Table -AutoSize -Property JobName, VMName, GuestProcessing, ProcessingMode, SqlLogHandling,
        @{ Name = 'LogEvery(min)'; Expression = { $_.SqlLogBackupEveryMin } }, SettingsFrom | Out-Host

    if (-not $OutputPath) {
        $folder = if ($outputFolder) { $outputFolder } elseif ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
        $fileName = 'VeeamGuestProcessing_{0}_{1:yyyyMMdd_HHmm}.csv' -f ($connectedServer -replace '[^\w.-]', '_'), (Get-Date)
        $OutputPath = Join-Path $folder $fileName
    }
    try {
        $result | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8BOM
    }
    catch {
        $fallback = Join-Path ([IO.Path]::GetTempPath()) ('{0:yyyyMMdd_HHmmss}_{1}' -f (Get-Date), (Split-Path $OutputPath -Leaf))
        Write-Warning "Could not write '$OutputPath': $($_.Exception.Message) Saving to '$fallback' instead."
        $OutputPath = $fallback
        $result | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8BOM
    }

    $failed = @($result | Where-Object VMName -eq '[ERROR - job not checked]').Count
    $notExpanded = @($result | Where-Object { $_.VMName.StartsWith('[not expanded] ') }).Count
    $jobCount = @($result | Select-Object -ExpandProperty JobName -Unique).Count
    Write-Host "Found $($result.Count - $failed - $notExpanded) VM(s) or agent job object(s) in $jobCount job(s)."
    if ($notExpanded) { Write-Warning "$notExpanded container(s) could not be expanded. See the rows marked [not expanded]." }
    if ($failed) { Write-Warning "$failed job(s) could not be checked. See the rows marked [ERROR - job not checked]." }
    Write-Host "CSV saved to: $OutputPath"

    if ($PassThru) { $result }
}
finally {
    if ($connectedByScript) {
        Disconnect-VBRServer
        Write-Host "Disconnected from '$connectedServer'."
    }
}
