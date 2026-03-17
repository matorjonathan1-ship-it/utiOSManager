# utiOSManager: Windows System Management Script

This PowerShell script (`utiOSManager.ps1`) is a comprehensive tool designed for managing Windows systems. It combines functionalities for activating Microsoft Windows and Office products using a Key Management Service (KMS) host with powerful disk management capabilities, including listing, initializing, partitioning, and formatting drives.

## Disclaimer

**This script is provided for educational and legitimate organizational use only.** It is crucial to understand and comply with Microsoft's licensing terms and conditions. Unauthorized use of KMS activation or activation against non-legitimate KMS servers is a violation of software licensing agreements. Furthermore, disk management operations are inherently destructive and can lead to data loss if not used carefully. The author and maintainers of this script are not responsible for any misuse, data loss, or legal consequences arising from its use.

## Features

*   **Windows Activation:** Configures and activates the Windows operating system against a specified KMS host.
*   **Office Activation:** Configures and activates installed Microsoft Office products (Office 2016, 2019, 2021) against a specified KMS host.
*   **Customizable KMS Host/Port:** Allows specifying a custom KMS server hostname/IP and port.
*   **Selective Activation:** Option to activate only Windows, only Office, or both.
*   **Disk Listing:** Lists all physical disks and their properties for easy identification.
*   **Disk Initialization:** Initializes a specified disk (e.g., to GPT) for new use.
*   **Partition Creation:** Creates primary partitions on a specified disk.
*   **Partition Formatting:** Formats partitions with specified file systems (e.g., NTFS, FAT32).
*   **Safety Confirmations:** Includes user confirmations before executing destructive disk operations.

## Requirements

*   **Administrative Privileges:** The script must be run with administrator rights.
*   **KMS Host (for activation):** A reachable and properly configured KMS server on your network.
*   **Windows 10/11 or Windows Server:** Disk management features are designed for modern Windows operating systems.

## How to Use

1.  **Download the script:** Save the `utiOSManager.ps1` file to your local machine.

2.  **Open PowerShell as Administrator:**
    *   Right-click the Start button and select "Windows PowerShell (Admin)" or "Command Prompt (Admin)".

3.  **Navigate to the script directory:**
    ```powershell
    cd "C:\Path\To\Your\Script"
    ```
    (Replace `C:\Path\To\Your\Script` with the actual path where you saved the script.)

4.  **Execute the script with desired parameters:**

    *   **Activate both Windows and Office using a specific KMS host:**
        ```powershell
        .\utiOSManager.ps1 -KmsHost "your_kms_server.your_domain" -ActivateWindows -ActivateOffice
        ```
        (Replace `your_kms_server.your_domain` with the actual hostname or IP address of your KMS server.)

    *   **List all physical disks:**
        ```powershell
        .\utiOSManager.ps1 -ListDisks
        ```

    *   **Initialize Disk 1, create a primary partition, and format it with NTFS (requires confirmation):**
        ```powershell
        .\utiOSManager.ps1 -DiskNumber 1 -InitializeDisk -CreatePartition -FormatPartition -FileSystem NTFS
        ```
        **WARNING:** This operation is destructive and will erase all data on the specified disk.

    *   **Activate only Windows and list disks:**
        ```powershell
        .\utiOSManager.ps1 -KmsHost "192.168.1.100" -ActivateWindows -ListDisks
        ```

## Script Parameters

*   `-KmsHost <string>`: Specifies the hostname or IP address of your KMS server. If not provided, the script will attempt to use auto-discovery or a previously configured KMS host.
*   `-KmsPort <int>`: Specifies the port number of the KMS server. The default is `1688`.
*   `-ActivateWindows`: A switch parameter. Include this to activate Windows.
*   `-ActivateOffice`: A switch parameter. Include this to activate Microsoft Office.
*   `-ListDisks`: A switch parameter to list all physical disks and their properties.
*   `-DiskNumber <int>`: Specifies the disk number to perform operations on (e.g., partitioning, formatting). Required for disk operations.
*   `-InitializeDisk`: A switch parameter to initialize the specified disk (e.g., to GPT). Requires `-DiskNumber`.
*   `-CreatePartition`: A switch parameter to create a primary partition on the specified disk. Requires `-DiskNumber`.
*   `-FormatPartition`: A switch parameter to format a partition on the specified disk. Requires `-DiskNumber` and `-PartitionNumber`.
*   `-PartitionNumber <int>`: Specifies the partition number to format. Required with `-FormatPartition`.
*   `-FileSystem <string>`: Specifies the file system for formatting (e.g., `NTFS`, `FAT32`). Default is `NTFS`.

## Troubleshooting

*   **"This script must be run with administrative privileges."**: Ensure you opened PowerShell as an administrator.
*   **KMS Activation Issues**: Verify the KMS host is correct, reachable, and the KMS service is running. Check network connectivity and firewall rules. For Office, ensure the correct version is installed and the KMS host supports it.
*   **Disk Operation Errors**: Ensure the `DiskNumber` is correct. Disk operations can fail if the disk is in use, write-protected, or if there are underlying hardware issues. Always back up important data before performing disk operations.
*   **Error codes from `slmgr.vbs`, `ospp.vbs`, or `Get-Disk`/`Initialize-Disk`/`New-Partition`/`Format-Volume` cmdlets**: These tools provide specific error codes or messages. Search Microsoft's documentation for details on these to diagnose further.

## Contributing

Feel free to fork this repository, submit pull requests, or open issues if you find bugs or have suggestions for improvements.

## License

This project is licensed under the MIT License - see the `LICENSE` file for details.
