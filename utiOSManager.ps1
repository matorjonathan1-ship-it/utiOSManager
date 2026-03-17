# utiOSManager.ps1
# This script provides comprehensive management for Windows systems, including KMS activation for Windows and Office,
# and disk management functionalities like listing, partitioning, and formatting drives.
# It is designed for legitimate use in environments where a KMS server is properly licensed and configured.

<#
.SYNOPSIS
    Comprehensive Windows system management tool for KMS activation and disk operations.

.DESCRIPTION
    This script combines functionalities for activating Windows and Microsoft Office via KMS
    with powerful disk management capabilities. It allows users to list disks, initialize them,
    create partitions (including EFI and primary partitions), and format volumes.
    Safety mechanisms are included for destructive disk operations.

.PARAMETER KmsHost
    Specifies the hostname or IP address of the KMS server. If not provided, the script
    will attempt to use auto-discovery or a default KMS host if one is set in the system.

.PARAMETER KmsPort
    Specifies the port number of the KMS server. Default is 1688.

.PARAMETER ActivateWindows
    A switch parameter to activate Windows. If omitted, Windows activation will be skipped.

.PARAMETER ActivateOffice
    A switch parameter to activate Microsoft Office. If omitted, Office activation will be skipped.

.PARAMETER ListDisks
    A switch parameter to list all physical disks and their properties.

.PARAMETER DiskNumber
    Specifies the disk number to perform operations on (e.g., partitioning, formatting).
    Required for disk operations.

.PARAMETER InitializeDisk
    A switch parameter to initialize the specified disk (e.g., to GPT).
    Requires -DiskNumber.

.PARAMETER CreatePartition
    A switch parameter to create a primary partition on the specified disk.
    Requires -DiskNumber.

.PARAMETER FormatPartition
    A switch parameter to format a partition on the specified disk.
    Requires -DiskNumber and -PartitionNumber.

.PARAMETER PartitionNumber
    Specifies the partition number to format. Required with -FormatPartition.

.PARAMETER FileSystem
    Specifies the file system for formatting (e.g., NTFS, FAT32). Default is NTFS.

.EXAMPLE
    .\utiOSManager.ps1 -KmsHost "kms.example.com" -ActivateWindows -ActivateOffice
    Activates Windows and Office using "kms.example.com" as the KMS host.

.EXAMPLE
    .\utiOSManager.ps1 -ListDisks
    Lists all physical disks on the system.

.EXAMPLE
    .\utiOSManager.ps1 -DiskNumber 1 -InitializeDisk -CreatePartition -FormatPartition -FileSystem NTFS
    Initializes Disk 1, creates a primary partition, and formats it with NTFS.

.NOTES
    Requires administrative privileges to run.
    Disk operations are destructive. Use with extreme caution.
#>

[CmdletBinding()]
param (
    [string]$KmsHost,
    [int]$KmsPort = 1688,
    [switch]$ActivateWindows,
    [switch]$ActivateOffice,
    [switch]$ListDisks,
    [int]$DiskNumber,
    [switch]$InitializeDisk,
    [switch]$CreatePartition,
    [switch]$FormatPartition,
    [int]$PartitionNumber,
    [string]$FileSystem = "NTFS"
)

function Set-KmsHost {
    param (
        [string]$HostName,
        [int]$Port
    )
    Write-Host "Setting KMS host to $HostName:$Port..." -ForegroundColor Cyan
    try {
        # Set KMS host for Windows
        & cscript.exe "$env:SystemRoot\system32\slmgr.vbs" /skms "$HostName`:$Port"
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Windows KMS host set successfully." -ForegroundColor Green
        } else {
            Write-Warning "Failed to set Windows KMS host. Error code: $LASTEXITCODE"
        }

        # Set KMS host for Office (if Office is installed and relevant paths exist)
        $officePaths = @(
            "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
            "${env:ProgramFiles}\Microsoft Office\Office16",
            "${env:ProgramFiles(x86)}\Microsoft Office\Office15",
            "${env:ProgramFiles}\Microsoft Office\Office15"
        )

        $officeKmsHostSet = $false
        foreach ($path in $officePaths) {
            $osppPath = Join-Path $path "ospp.vbs"
            if (Test-Path $osppPath) {
                Write-Host "Attempting to set Office KMS host using $osppPath" -ForegroundColor DarkCyan
                & cscript.exe $osppPath /sethst:$HostName
                & cscript.exe $osppPath /setprt:$Port
                if ($LASTEXITCODE -eq 0) {
                    Write-Host "Office KMS host set successfully for $path." -ForegroundColor Green
                    $officeKmsHostSet = $true
                    break
                } else {
                    Write-Warning "Failed to set Office KMS host for $path. Error code: $LASTEXITCODE"
                }
            }
        }
        if (-not $officeKmsHostSet) {
            Write-Warning "Could not find a suitable Office installation to set KMS host, or failed to set it."
        }

    } catch {
        Write-Error "An error occurred while setting KMS host: $($_.Exception.Message)"
    }
}

function Activate-Product {
    param (
        [string]$ProductType
    )
    Write-Host "Attempting to activate $ProductType..." -ForegroundColor Cyan
    try {
        if ($ProductType -eq "Windows") {
            & cscript.exe "$env:SystemRoot\system32\slmgr.vbs" /ato
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Windows activated successfully." -ForegroundColor Green
            } else {
                Write-Warning "Failed to activate Windows. Error code: $LASTEXITCODE"
            }
        } elseif ($ProductType -eq "Office") {
            $officePaths = @(
                "${env:ProgramFiles(x86)}\Microsoft Office\Office16",
                "${env:ProgramFiles}\Microsoft Office\Office16",
                "${env:ProgramFiles(x86)}\Microsoft Office\Office15",
                "${env:ProgramFiles}\Microsoft Office\Office15"
            )

            $officeActivated = $false
            foreach ($path in $officePaths) {
                $osppPath = Join-Path $path "ospp.vbs"
                if (Test-Path $osppPath) {
                    Write-Host "Attempting to activate Office using $osppPath" -ForegroundColor DarkCyan
                    & cscript.exe $osppPath /act
                    if ($LASTEXITCODE -eq 0) {
                        Write-Host "Office activated successfully for $path." -ForegroundColor Green
                        $officeActivated = $true
                        break
                    } else {
                        Write-Warning "Failed to activate Office for $path. Error code: $LASTEXITCODE"
                    }
                }
            }
            if (-not $officeActivated) {
                Write-Warning "Could not find a suitable Office installation to activate, or failed to activate it."
            }
        }
    } catch {
        Write-Error "An error occurred during $ProductType activation: $($_.Exception.Message)"
    }
}

function Get-Confirmation {
    param (
        [string]$Message
    )
    $choice = Read-Host "$Message (Y/N)"
    return ($choice -eq "Y" -or $choice -eq "y")
}

function List-PhysicalDisks {
    Write-Host "Listing physical disks..." -ForegroundColor Cyan
    try {
        Get-Disk | Format-Table -AutoSize
    } catch {
        Write-Error "An error occurred while listing disks: $($_.Exception.Message)"
    }
}

function Initialize-AndPartitionDisk {
    param (
        [int]$DiskNumberToOperate,
        [switch]$CreatePrimaryPartition,
        [switch]$FormatNewPartition,
        [string]$FileSystemType
    )

    Write-Host "Attempting to initialize and partition Disk $DiskNumberToOperate..." -ForegroundColor Cyan

    if (-not (Get-Confirmation -Message "WARNING: All data on Disk $DiskNumberToOperate will be lost. Do you want to proceed?")) {
        Write-Warning "Disk initialization and partitioning cancelled by user."
        return
    }

    try {
        # Clear and Initialize Disk (GPT)
        Write-Host "Clearing and initializing Disk $DiskNumberToOperate as GPT..." -ForegroundColor DarkCyan
        Clear-Disk -Number $DiskNumberToOperate -RemoveData -RemoveOEM -Confirm:$false
        Initialize-Disk -Number $DiskNumberToOperate -PartitionStyle GPT -Confirm:$false

        # Create EFI System Partition (ESP)
        Write-Host "Creating EFI System Partition on Disk $DiskNumberToOperate..." -ForegroundColor DarkCyan
        New-Partition -DiskNumber $DiskNumberToOperate -UseMaximumSize -IsBoot -AssignDriveLetter:$false | Format-Volume -FileSystem FAT32 -NewFileSystemLabel "EFI" -Confirm:$false

        # Create Primary Partition
        if ($CreatePrimaryPartition) {
            Write-Host "Creating Primary Partition on Disk $DiskNumberToOperate..." -ForegroundColor DarkCyan
            $newPartition = New-Partition -DiskNumber $DiskNumberToOperate -UseMaximumSize
            
            if ($FormatNewPartition) {
                Write-Host "Formatting new partition on Disk $DiskNumberToOperate with $FileSystemType..." -ForegroundColor DarkCyan
                $newPartition | Format-Volume -FileSystem $FileSystemType -NewFileSystemLabel "Primary" -Confirm:$false
            }
        }
        Write-Host "Disk $DiskNumberToOperate operations completed successfully." -ForegroundColor Green

    } catch {
        Write-Error "An error occurred during disk operations on Disk $DiskNumberToOperate: $($_.Exception.Message)"
    }
}

# Main script logic
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run with administrative privileges."
    exit 1
}

# Disk Management Operations
if ($ListDisks) {
    List-PhysicalDisks
}

if ($PSBoundParameters.ContainsKey('DiskNumber')) {
    if ($InitializeDisk -or $CreatePartition -or $FormatPartition) {
        Initialize-AndPartitionDisk -DiskNumber $DiskNumber -CreatePrimaryPartition:$CreatePartition -FormatNewPartition:$FormatPartition -FileSystemType $FileSystem
    }
}

# KMS Activation Operations
if ($PSBoundParameters.ContainsKey('KmsHost')) {
    Set-KmsHost -HostName $KmsHost -Port $KmsPort
}

if ($ActivateWindows) {
    Activate-Product -ProductType "Windows"
}

if ($ActivateOffice) {
    Activate-Product -ProductType "Office"
}

if (-not $ActivateWindows -and -not $ActivateOffice -and -not $ListDisks -and -not $InitializeDisk -and -not $CreatePartition -and -not $FormatPartition) {
    Write-Host "No specific operation specified. Please use -ActivateWindows, -ActivateOffice, -ListDisks, or disk management parameters." -ForegroundColor Yellow
    Write-Host "Attempting to activate both Windows and Office by default if KMS host is provided..." -ForegroundColor Yellow
    if ($PSBoundParameters.ContainsKey('KmsHost')) {
        Activate-Product -ProductType "Windows"
        Activate-Product -ProductType "Office"
    } else {
        Write-Warning "No KMS host provided for default activation. Skipping activation."
    }
}

Write-Host "Script execution completed." -ForegroundColor Green
