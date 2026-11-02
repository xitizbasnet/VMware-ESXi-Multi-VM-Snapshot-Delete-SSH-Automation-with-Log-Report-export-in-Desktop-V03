# ==============================================================================
# VMware ESXi VM + Snapshot Bulk Delete & Excel Reporting (Verified SSH Edition)
# - Parallel dispatch via detached background vim-cmd execution
# - Hard verification against live snapshot trees via `snapshot.get`
# - Retries and waits until snapshot ID is physically removed from VM tree
# - Styled Blue & White table report with cell borders, no sheet gridlines
# - Auto-launches Excel report upon completion
# ==============================================================================

Clear-Host
$ErrorActionPreference = "Stop"

# ==============================================================================
# CONFIGURATION
# ==============================================================================

$ESXiIP   = "192.168.5.2"
$Username = "root"

$DesktopPath  = [Environment]::GetFolderPath("Desktop")
$LogDirectory = Join-Path $DesktopPath "VMware Automation Deletion Log"

# Full month timestamp: 2026-October-29_08-14-52
$DateStamp    = Get-Date -Format "MMMM-dd-yyyy_HH-mm-ss"
$LogFile      = Join-Path $LogDirectory "ESXi_${ESXiIP}_Snapshot_Delete_Report_$DateStamp.xlsx"

$Global:SessionLogs = [System.Collections.Generic.List[PSCustomObject]]::new()

# ==============================================================================
# 1. INITIALIZE DIRECTORY
# ==============================================================================

try {
    if (-not (Test-Path -LiteralPath $LogDirectory)) {
        New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction Stop | Out-Null
    }
}
catch {
    Write-Host "ERROR: Unable to create log directory: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press ENTER to exit"
    exit 1
}

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "   ESXi VM PARALLEL BULK SNAPSHOT PURGE TOOL (VERIFIED)" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "ESXi Host     : $ESXiIP"
Write-Host "Username      : $Username"
Write-Host "Desktop Log   : $LogFile"
Write-Host ""

# ==============================================================================
# 2. PREREQUISITES & MODULE SETUP
# ==============================================================================

Write-Host "[1/6] Checking PowerShell Version..." -ForegroundColor Yellow
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Host "ERROR: PowerShell 5.1 or newer is required." -ForegroundColor Red
    Read-Host "Press ENTER to exit"
    exit 1
}

Write-Host "[2/6] Checking VMware PowerCLI Core..." -ForegroundColor Yellow
$CoreModule = Get-Module -ListAvailable -Name VMware.VimAutomation.Core | Sort-Object Version -Descending | Select-Object -First 1
if (-not $CoreModule) {
    Write-Host "Installing VMware PowerCLI..." -ForegroundColor Yellow
    Install-Module VMware.PowerCLI -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}
Import-Module VMware.VimAutomation.Core -ErrorAction Stop

Write-Host "[3/6] Checking Posh-SSH..." -ForegroundColor Yellow
$PoshSSH = Get-Module -ListAvailable -Name Posh-SSH | Sort-Object Version -Descending | Select-Object -First 1
if (-not $PoshSSH) {
    Write-Host "Installing Posh-SSH module..." -ForegroundColor Yellow
    Install-Module -Name Posh-SSH -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}
Import-Module Posh-SSH -ErrorAction Stop

Write-Host "[4/6] Checking ImportExcel..." -ForegroundColor Yellow
$ImportExcelModule = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
if (-not $ImportExcelModule) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
}
Import-Module ImportExcel -ErrorAction Stop

Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -Scope User -ErrorAction SilentlyContinue | Out-Null
Write-Host "All prerequisites verified." -ForegroundColor Green
Write-Host ""

# ==============================================================================
# 3. AUTHENTICATION & CONNECTION RETRY LOOP
# ==============================================================================

$PowerCLIConnection = $null
$SSHSession         = $null
$Credential         = $null

while ($null -eq $PowerCLIConnection -or $null -eq $SSHSession) {
    Write-Host "[5/6] Authentication" -ForegroundColor Yellow
    $Password = Read-Host "Enter password for $Username" -AsSecureString

    if (-not $Password) {
        Write-Host "Password cannot be empty. Please try again.`n" -ForegroundColor Red
        continue
    }

    $Credential = New-Object System.Management.Automation.PSCredential($Username, $Password)

    Write-Host ""
    Write-Host "Testing PowerCLI & SSH connection to $ESXiIP..." -ForegroundColor Cyan

    try {
        $PowerCLIConnection = Connect-VIServer -Server $ESXiIP -Credential $Credential -ErrorAction Stop
        $SSHSession = New-SSHSession -ComputerName $ESXiIP -Credential $Credential -AcceptKey -ErrorAction Stop
        Write-Host "SUCCESS: Connected via PowerCLI and SSH." -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host ""
        Write-Host "============================================================" -ForegroundColor Red
        Write-Host "           LOGIN FAILED: INCORRECT PASSWORD OR REJECTED" -ForegroundColor Red
        Write-Host "============================================================" -ForegroundColor Red
        Write-Host "Error Details: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "You entered an incorrect password or SSH service is stopped. Please re-enter." -ForegroundColor Yellow
        Write-Host "------------------------------------------------------------`n"

        if ($PowerCLIConnection) {
            Disconnect-VIServer -Server $PowerCLIConnection -Confirm:$false -ErrorAction SilentlyContinue
        }
        if ($SSHSession) {
            Remove-SSHSession -SessionId $SSHSession.SessionId -ErrorAction SilentlyContinue
        }

        $PowerCLIConnection = $null
        $SSHSession         = $null
    }
}

# ==============================================================================
# 4. GATHER HOST & VM DETAILS
# ==============================================================================

Write-Host "[6/6] Collecting ESXi Host and VM inventory..." -ForegroundColor Yellow
$ESXi = Get-VMHost -Server $PowerCLIConnection -ErrorAction Stop
$VMs  = @(Get-VM -Location $ESXi -ErrorAction SilentlyContinue)

Write-Host "Host Name : $($ESXi.Name)"
Write-Host "Version   : $($ESXi.Version) (Build $($ESXi.Build))"
Write-Host "Total VMs : $($VMs.Count)"
Write-Host ""

# ==============================================================================
# 5. LOGGING & EXCEL FORMATTING FUNCTIONS
# ==============================================================================

function Add-SessionLog {
    param (
        [string]$VMName,
        [string]$SnapshotNumber,
        [string]$SnapshotName,
        [string]$SnapshotId,
        [string]$Created,
        [string]$SizeGB,
        [string]$Quiesced,
        [string]$PowerState,
        [string]$Action,
        [string]$Result,
        [string]$Details
    )

    $LogObject = [PSCustomObject]@{
        Timestamp      = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        ESXiIP         = $ESXiIP
        ESXiHost       = $ESXi.Name
        Username       = $Username
        VMName         = $VMName
        SnapshotNumber = $SnapshotNumber
        SnapshotName   = $SnapshotName
        SnapshotID     = $SnapshotId
        Created        = $Created
        SizeGB         = $SizeGB
        Quiesced       = $Quiesced
        PowerState     = $PowerState
        Action         = $Action
        Result         = $Result
        Details        = $Details
    }

    $Global:SessionLogs.Add($LogObject)
}

function Export-StyledExcelReport {
    if ($Global:SessionLogs.Count -eq 0) {
        Write-Host "No actions were performed. Skipping Excel report." -ForegroundColor Yellow
        return
    }

    Write-Host "Formatting Excel report with Blue/White styling..." -ForegroundColor Cyan

    try {
        $excelPackage = $Global:SessionLogs | Export-Excel -Path $LogFile `
            -WorksheetName "Snapshot Log" `
            -TableName "SnapshotAuditTable" `
            -TableStyle Medium2 `
            -AutoSize `
            -PassThru

        $ws = $excelPackage.Workbook.Worksheets["Snapshot Log"]

        # 1. Turn off background sheet gridlines
        $ws.View.ShowGridLines = $false

        # 2. Add border styling inside table cells
        $tableRange = $ws.Dimension.Address
        $cells = $ws.Cells[$tableRange]

        $cells.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $cells.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $cells.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $cells.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

        Close-ExcelPackage $excelPackage -Show:$false
        Write-Host "Excel report saved: $LogFile" -ForegroundColor Green

        # 3. Automatically launch Excel
        if (Test-Path -LiteralPath $LogFile) {
            Write-Host "Opening report in Excel..." -ForegroundColor Cyan
            Invoke-Item $LogFile
        }
    }
    catch {
        Write-Host "Warning during Excel formatting: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

# ==============================================================================
# 6. SNAPSHOT INVENTORY SCANNER
# ==============================================================================

function Get-AllSnapshotInventory {
    $Inventory = @()

    foreach ($VM in $VMs) {
        try {
            $VMSnapshots = @(Get-Snapshot -VM $VM -ErrorAction SilentlyContinue | Sort-Object Created)
        }
        catch {
            continue
        }

        $NumericVMId = $VM.Id -replace '\D+'

        $Position = 1
        foreach ($Snapshot in $VMSnapshots) {
            $NumericSnapId = ($Snapshot.Id -split '-')[-1]

            $Inventory += [PSCustomObject]@{
                VMName         = $VM.Name
                VMObject       = $VM
                NumericVMId    = $NumericVMId
                SnapshotNumber = $Position
                SnapshotObject = $Snapshot
                SnapshotName   = $Snapshot.Name
                SnapshotId     = $Snapshot.Id
                NumericSnapId  = $NumericSnapId
                Created        = $Snapshot.Created
                SizeGB         = $Snapshot.SizeGB
                Quiesced       = $Snapshot.Quiesced
                PowerState     = $Snapshot.PowerState
            }
            $Position++
        }
    }

    return $Inventory
}

function Show-SnapshotInventory {
    param ([array]$Inventory)

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "                ALL VM SNAPSHOT INFORMATION" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""

    if ($Inventory.Count -eq 0) {
        Write-Host "NO SNAPSHOTS FOUND ON ANY VM." -ForegroundColor Green
        Write-Host ""
        return
    }

    $MaxPosition = ($Inventory | Measure-Object SnapshotNumber -Maximum).Maximum

    for ($Position = 1; $Position -le $MaxPosition; $Position++) {
        $PositionSnapshots = @($Inventory | Where-Object { $_.SnapshotNumber -eq $Position })

        Write-Host "************************************************************" -ForegroundColor Yellow
        Write-Host "                     SNAPSHOT #$Position" -ForegroundColor Yellow
        Write-Host "************************************************************" -ForegroundColor Yellow
        Write-Host ""

        if ($PositionSnapshots.Count -eq 0) {
            Write-Host "No VM has Snapshot #$Position."
            continue
        }

        foreach ($Item in $PositionSnapshots) {
            Write-Host "VM          : $($Item.VMName) (VMID: $($Item.NumericVMId))"
            Write-Host "Snapshot    : $($Item.SnapshotName) (SnapID: $($Item.NumericSnapId))"
            Write-Host "Created     : $($Item.Created)"
            Write-Host "Size        : $([math]::Round($Item.SizeGB,2)) GB"
            Write-Host "Quiesced    : $($Item.Quiesced)"
            Write-Host "Power State : $($Item.PowerState)"
            Write-Host "------------------------------------------------------------"
        }

        Write-Host "Total VMs with Snapshot #$Position : $($PositionSnapshots.Count)" -ForegroundColor Cyan
        Write-Host ""
    }
}

# ==============================================================================
# 7. PARALLEL DELETION ENGINE WITH LIVE REVERIFICATION
# ==============================================================================

function Remove-SnapshotPositionSSH {
    param (
        [int]$SnapshotPosition,
        [array]$Inventory
    )

    $Targets = @($Inventory | Where-Object { $_.SnapshotNumber -eq $SnapshotPosition })

    if ($Targets.Count -eq 0) {
        Write-Host "No VMs have Snapshot #$SnapshotPosition." -ForegroundColor Yellow
        return
    }

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "         PARALLEL BULK SNAPSHOT DELETE (ALL AT ONCE)" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "Snapshot Position : #$SnapshotPosition"
    Write-Host "VMs affected      : $($Targets.Count)"
    Write-Host ""

    foreach ($Target in $Targets) {
        Write-Host "VM       : $($Target.VMName) [ID: $($Target.NumericVMId)]"
        Write-Host "Snapshot : $($Target.SnapshotName) [SnapID: $($Target.NumericSnapId)]"
        Write-Host "Size     : $([math]::Round($Target.SizeGB,2)) GB"
        Write-Host "------------------------------------------------------------"
    }

    Write-Host ""
    Write-Host "Type DELETE to launch deletion on ALL $($Targets.Count) VMs simultaneously." -ForegroundColor Red
    $Confirmation = Read-Host "Type DELETE to confirm"

    if ($Confirmation -cne "DELETE") {
        Write-Host "Operation cancelled." -ForegroundColor Green
        return
    }

    Write-Host ""
    Write-Host "Dispatching background deletion to ALL $($Targets.Count) VMs simultaneously..." -ForegroundColor Cyan

    # Phase 1: Launch background delete commands on all target VMs
    foreach ($Target in $Targets) {
        Write-Host "-> Dispatched: [$($Target.VMName)] (VMID: $($Target.NumericVMId))" -ForegroundColor Yellow
        $BgCommand = "nohup vim-cmd vmsvc/snapshot.remove $($Target.NumericVMId) $($Target.NumericSnapId) 0 > /dev/null 2>&1 &"
        Invoke-SSHCommand -SessionId $SSHSession.SessionId -Command $BgCommand | Out-Null
    }

    Write-Host ""
    Write-Host "All consolidation jobs launched in background." -ForegroundColor Green
    Write-Host "Entering live verification loop (monitoring snapshot tree for each VM)..." -ForegroundColor Cyan

    # Phase 2: Live reverification polling loop
    $PendingTargets = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($t in $Targets) { $PendingTargets.Add($t) }

    $PollIntervalSec = 5
    $MaxWaitMinutes  = 30
    $StartTime       = Get-Date

    while ($PendingTargets.Count -gt 0) {
        $Elapsed = ((Get-Date) - $StartTime).TotalMinutes
        if ($Elapsed -gt $MaxWaitMinutes) {
            Write-Host "TIMEOUT WARNING: Max wait time ($MaxWaitMinutes mins) reached." -ForegroundColor Red
            break
        }

        Write-Host "Verifying status... ($($PendingTargets.Count) VM(s) still merging delta disks)" -ForegroundColor Cyan
        
        $StillPending = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($Target in $PendingTargets) {
            # Query ESXi's live snapshot hierarchy directly for this VM
            $SnapCheck = Invoke-SSHCommand -SessionId $SSHSession.SessionId -Command "vim-cmd vmsvc/snapshot.get $($Target.NumericVMId)"
            $Output = $SnapCheck.Output

            # Check if Snapshot ID or Snapshot Name is still present in the tree output
            $IdPattern = "Id\s*:\s*$($Target.NumericSnapId)\b"
            $SnapshotStillPresent = ($Output -match $IdPattern) -or ($Output -match [regex]::Escape($Target.SnapshotName))

            if ($SnapshotStillPresent) {
                # Still merging
                $StillPending.Add($Target)
            }
            else {
                # Physically verified gone from the VM snapshot tree
                Write-Host "VERIFIED DELETED: [$($Target.VMName)] -> $($Target.SnapshotName)" -ForegroundColor Green

                Add-SessionLog -VMName $Target.VMName -SnapshotNumber $Target.SnapshotNumber `
                    -SnapshotName $Target.SnapshotName -SnapshotId $Target.SnapshotId -Created $Target.Created `
                    -SizeGB $Target.SizeGB -Quiesced $Target.Quiesced -PowerState $Target.PowerState `
                    -Action "DELETE SNAPSHOT #$SnapshotPosition (PARALLEL)" -Result "SUCCESS" `
                    -Details "Verified absent from live VM snapshot tree"
            }
        }

        $PendingTargets = $StillPending

        if ($PendingTargets.Count -gt 0) {
            Start-Sleep -Seconds $PollIntervalSec
        }
    }

    # Handle any that timed out
    if ($PendingTargets.Count -gt 0) {
        foreach ($TimedOut in $PendingTargets) {
            Write-Host "NOT CONFIRMED: [$($TimedOut.VMName)] is still consolidating." -ForegroundColor Yellow
            Add-SessionLog -VMName $TimedOut.VMName -SnapshotNumber $TimedOut.SnapshotNumber `
                -SnapshotName $TimedOut.SnapshotName -SnapshotId $TimedOut.SnapshotId -Created $TimedOut.Created `
                -SizeGB $TimedOut.SizeGB -Quiesced $TimedOut.Quiesced -PowerState $TimedOut.PowerState `
                -Action "DELETE SNAPSHOT #$SnapshotPosition (PARALLEL)" -Result "IN PROGRESS" `
                -Details "Task still active on ESXi after $MaxWaitMinutes minutes"
        }
    }

    Write-Host ""
    Write-Host "Reverification complete for Snapshot #$SnapshotPosition." -ForegroundColor Green
}

function Remove-AllSnapshotsSSH {
    param ([array]$Inventory)

    if ($Inventory.Count -eq 0) { return }

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "      DELETE ALL SNAPSHOTS FROM ALL VMs AT ONCE (PARALLEL)" -ForegroundColor Red
    Write-Host "============================================================" -ForegroundColor Red
    Write-Host "Total snapshots to remove: $($Inventory.Count)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Type DELETE ALL to confirm clearing every snapshot." -ForegroundColor Red
    $Confirmation = Read-Host "Type DELETE ALL to confirm"

    if ($Confirmation -cne "DELETE ALL") {
        Write-Host "Cancelled." -ForegroundColor Green
        return
    }

    $UniqueVMs = $Inventory | Select-Object -Property VMName, NumericVMId -Unique

    Write-Host ""
    Write-Host "Dispatching parallel removeall commands..." -ForegroundColor Cyan

    foreach ($Item in $UniqueVMs) {
        Write-Host "-> Dispatched: [$($Item.VMName)] (VMID: $($Item.NumericVMId))" -ForegroundColor Yellow
        $BgCommand = "nohup vim-cmd vmsvc/snapshot.removeall $($Item.NumericVMId) > /dev/null 2>&1 &"
        Invoke-SSHCommand -SessionId $SSHSession.SessionId -Command $BgCommand | Out-Null
    }

    Write-Host ""
    Write-Host "Jobs launched. Verifying complete snapshot removal per VM..." -ForegroundColor Cyan

    $PendingVMs = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($vm in $UniqueVMs) { $PendingVMs.Add($vm) }

    $PollIntervalSec = 5
    $MaxWaitMinutes  = 30
    $StartTime       = Get-Date

    while ($PendingVMs.Count -gt 0) {
        $Elapsed = ((Get-Date) - $StartTime).TotalMinutes
        if ($Elapsed -gt $MaxWaitMinutes) {
            Write-Host "TIMEOUT WARNING: Max wait time ($MaxWaitMinutes mins) reached." -ForegroundColor Red
            break
        }

        Write-Host "Verifying status... ($($PendingVMs.Count) VM(s) still merging delta disks)" -ForegroundColor Cyan
        $StillPending = [System.Collections.Generic.List[PSCustomObject]]::new()

        foreach ($Item in $PendingVMs) {
            $SnapCheck = Invoke-SSHCommand -SessionId $SSHSession.SessionId -Command "vim-cmd vmsvc/snapshot.get $($Item.NumericVMId)"
            $Output = $SnapCheck.Output

            # If there are no snapshots, ESXi prints "Get Snapshot failed: This virtual machine does not have any snapshots" or similar
            if ($Output -match "does not have any snapshots" -or $Output -notmatch "Snapshot Id") {
                Write-Host "VERIFIED: All snapshots cleared on [$($Item.VMName)]" -ForegroundColor Green

                $VMRecords = $Inventory | Where-Object { $_.NumericVMId -eq $Item.NumericVMId }
                foreach ($rec in $VMRecords) {
                    Add-SessionLog -VMName $rec.VMName -SnapshotNumber $rec.SnapshotNumber `
                        -SnapshotName $rec.SnapshotName -SnapshotId $rec.SnapshotId -Created $rec.Created `
                        -SizeGB $rec.SizeGB -Quiesced $rec.Quiesced -PowerState $rec.PowerState `
                        -Action "DELETE ALL (PARALLEL)" -Result "SUCCESS" `
                        -Details "Verified empty snapshot tree via snapshot.get"
                }
            }
            else {
                $StillPending.Add($Item)
            }
        }

        $PendingVMs = $StillPending

        if ($PendingVMs.Count -gt 0) {
            Start-Sleep -Seconds $PollIntervalSec
        }
    }

    if ($PendingVMs.Count -gt 0) {
        foreach ($Item in $PendingVMs) {
            Write-Host "NOT CONFIRMED: [$($Item.VMName)] still processing." -ForegroundColor Yellow
            $VMRecords = $Inventory | Where-Object { $_.NumericVMId -eq $Item.NumericVMId }
            foreach ($rec in $VMRecords) {
                Add-SessionLog -VMName $rec.VMName -SnapshotNumber $rec.SnapshotNumber `
                    -SnapshotName $rec.SnapshotName -SnapshotId $rec.SnapshotId -Created $rec.Created `
                    -SizeGB $rec.SizeGB -Quiesced $rec.Quiesced -PowerState $rec.PowerState `
                    -Action "DELETE ALL (PARALLEL)" -Result "IN PROGRESS" `
                    -Details "Consolidation still in progress after $MaxWaitMinutes minutes"
            }
        }
    }

    Write-Host ""
    Write-Host "Bulk removal reverification finished." -ForegroundColor Green
}

# ==============================================================================
# 8. INTERACTIVE MENU LOOP
# ==============================================================================

function Start-SnapshotManagement {
    while ($true) {
        $SnapshotInventory = @(Get-AllSnapshotInventory)
        Show-SnapshotInventory -Inventory $SnapshotInventory

        if ($SnapshotInventory.Count -eq 0) {
            Write-Host "============================================================" -ForegroundColor Green
            Write-Host "           THERE ARE NO SNAPSHOTS REMAINING" -ForegroundColor Green
            Write-Host "============================================================" -ForegroundColor Green
            Write-Host ""
            Write-Host "[0] Exit"
            Write-Host ""
            $NoSnapChoice = Read-Host "Select option"
            if ($NoSnapChoice -eq "0") { return }
            continue
        }

        $MaxPosition = ($SnapshotInventory | Measure-Object SnapshotNumber -Maximum).Maximum

        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host "                   SNAPSHOT DELETE MENU" -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "[A] Delete ALL snapshots from ALL VMs at once (Parallel)" -ForegroundColor Red

        for ($Position = 1; $Position -le $MaxPosition; $Position++) {
            $Count = @($SnapshotInventory | Where-Object { $_.SnapshotNumber -eq $Position }).Count
            Write-Host "[$Position] Delete Snapshot #$Position from ALL $Count VM(s) at once"
        }

        Write-Host ""
        Write-Host "[R] Refresh snapshot list"
        Write-Host "[0] Exit"
        Write-Host ""

        $DeleteChoice = Read-Host "Select option"

        if ($DeleteChoice -eq "0") {
            Write-Host "Exiting snapshot management." -ForegroundColor Yellow
            return
        }

        if ($DeleteChoice -match "^[Rr]$") {
            Write-Host "Refreshing inventory..." -ForegroundColor Cyan
            continue
        }

        if ($DeleteChoice -match "^[Aa]$") {
            Remove-AllSnapshotsSSH -Inventory $SnapshotInventory
            Write-Host "Refreshing snapshot list..." -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            continue
        }

        $SelectedPosition = 0
        $IsNumber = [int]::TryParse($DeleteChoice, [ref]$SelectedPosition)

        if ($IsNumber -and $SelectedPosition -ge 1 -and $SelectedPosition -le $MaxPosition) {
            Remove-SnapshotPositionSSH -SnapshotPosition $SelectedPosition -Inventory $SnapshotInventory
            Write-Host ""
            Write-Host "Rescanning snapshots after deletion..." -ForegroundColor Cyan
            Start-Sleep -Seconds 2
            continue
        }
        else {
            Write-Host "Invalid selection. Please try again." -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}

# ==============================================================================
# EXECUTION & SHUTDOWN
# ==============================================================================

Start-SnapshotManagement

Export-StyledExcelReport

$FinalInventory = @(Get-AllSnapshotInventory)

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "                          SUMMARY" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "ESXi Host           : $($ESXi.Name)"
Write-Host "Remaining Snapshots : $($FinalInventory.Count)"
Write-Host "Excel Audit Log     : $LogFile"
Write-Host ""

if ($SSHSession) {
    Remove-SSHSession -SessionId $SSHSession.SessionId -ErrorAction SilentlyContinue | Out-Null
}
if ($PowerCLIConnection) {
    Disconnect-VIServer -Server $PowerCLIConnection -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
}

Write-Host "Disconnected from ESXi." -ForegroundColor Green
Write-Host ""
Read-Host "Press ENTER to close"
