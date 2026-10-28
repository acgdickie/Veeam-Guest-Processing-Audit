# Veeam Guest Processing Audit

A PowerShell script for **Veeam Backup & Replication v13**.

It lists every VM that Veeam backs up with **guest processing** (application-aware processing) turned on.
For each VM it shows the job, and the **SQL Server log settings** that Veeam uses for that VM.

The script **only reads** data. It does not change any job or setting.

## What it shows

One row for each VM in each job. A VM is listed when:

- the job has application-aware processing turned on, and
- the VM is not excluded from the job (or by a global exclusion), and
- application processing is not turned off for that VM.

It checks:

- VM backup jobs, replication jobs and VM copy jobs (VMware vSphere and Hyper-V)
- Veeam Agent jobs

It finds the VMs in each job, also when they are inside a container. A container is a group of VMs, such as a vCenter, cluster, host, folder, tag or Hyper-V VM group.

## What you need

- **PowerShell 7.** Veeam 13.1 needs **PowerShell 7.6 or later**.
- The **Veeam Backup & Replication v13 console** on the machine that runs the script. The console build must match the backup server build.
- A Veeam account that can read the jobs.

## How to run it

Open a **new** PowerShell 7 window. Then run one of these:

```powershell
# On the Veeam backup server
.\Get-VeeamGuestProcessingAudit.ps1

# From another machine
.\Get-VeeamGuestProcessingAudit.ps1 -Server vbr01 -Credential (Get-Credential)

# Also show VMs where guest processing is off, and check which VMs have SQL Server
.\Get-VeeamGuestProcessingAudit.ps1 -Server vbr01 -IncludeDisabled -DetectSql
```

The result shows on the screen. It is also saved to a CSV file in the script folder.

For all the options, run:

```powershell
Get-Help .\Get-VeeamGuestProcessingAudit.ps1 -Full
```

## Options

| Option | What it does |
|---|---|
| `-Server` | The Veeam backup server. Default: this machine. |
| `-Credential` | The account to sign in with. Default: your Windows account. |
| `-Port` | The port of the backup server. Default: the Veeam default. |
| `-ForceAcceptTlsCertificate` | Accepts the server certificate without a prompt. |
| `-JobName` | Only check these jobs. Wildcards work, for example `'SQL*'`. |
| `-OutputPath` | Where to save the CSV. A file or a folder. |
| `-IncludeDisabled` | Also show VMs where guest processing is off. |
| `-DetectSql` | Adds a SqlFound column (Yes/No). This can take a long time on a big server. |
| `-PassThru` | Also sends the results down the pipeline. |

## The SQL log settings

The `SqlLogHandling` column shows one of these:

| Value | Meaning |
|---|---|
| Truncate logs | Veeam truncates the SQL logs after a good backup. |
| Do not truncate logs | Veeam does not touch the SQL logs. |
| Back up logs (log shipping) | Veeam backs up the SQL logs. `SqlLogBackupEveryMin` and `SqlLogRetention` show how often and for how long. |
| Copy only | A copy-only backup. The SQL logs are not touched. |
| N/A (Linux guest) | SQL Server processing is for Windows only. |

## Read the SettingsFrom column

`SettingsFrom` tells you where the settings for that VM came from. For example: the VM itself, a tag, a folder or a host.

Check the rows that show:

- **(assumed)**: the script could not match the VM to one job object. All the possible objects have the same settings, so the values are most likely correct.
- **Unknown**: the script could not match the VM, and the possible objects have different settings. Check this VM in the Veeam console.

Also read the warnings on the screen. They tell you when part of the inventory could not be read, for example when a vCenter is offline.

## Good to know

- The script picks the settings the same way Veeam does: the most specific object wins. For example, a VM entry wins over a tag, and a tag wins over a folder.
- `InLastSession` and `LastResult` come from the latest job run.
- With `-DetectSql`, a VM is matched to SQL Server restore points by name only.
- CDP policies and backup copy jobs are not checked. The script tells you about them.
- Check a few VMs in the Veeam console the first time you run it.

## Licence

MIT. See [LICENSE](LICENSE).

## Disclaimer

Use this script at your own risk. Test it before you use it on a production server. This script is not made or supported by Veeam.
