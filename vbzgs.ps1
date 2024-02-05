param(
    [string]$CustomOutputDirectory
)

# Default output directory if $CustomOutputDirectory is not specified
if (-not $CustomOutputDirectory) {
    $CustomOutputDirectory = "C:\VBZ"
}

# Specify the paths of the files you want to check along with their display names
$filePaths = @{
	"C:\VBZ\Albireo Beheer\Albireo Beheer.exe" = "Albireo Beheer"    
	"C:\VBZ\tools\Albireo IP Debug.exe" = "Albireo IP Debugger"
    "C:\VBZ\INEX500\inex500.exe" = "INEX500"
	"C:\VBZ\INEX500_Backup\InexBackup.exe" = "INEX Backup"    
    "C:\VBZ\INEX500ServiceTool\inex500servicetool.exe" = "INEX500 ServiceTool"
	"C:\VBZ\DWS\DWS.EXE" = "DWS"
    "C:\VBZ\tools\PoEWatchdog\PoEWatchdog.exe" = "PoEWatchdog"
    "C:\VBZ\tools\PoEWatchdogNET\PoE Watchdog NET.exe" = "PoE Watchdog.NET"
    "C:\VBZ\Vios2\MQTT_V2EService.exe" = "V2E MQTT Bridge"
    "C:\VBZ\Vios2\V2E.exe" = "ViosEngine"
    "C:\VBZ\Zorg Management\ZORGMANA.exe" = "Zorgmanager"
    # Add more file paths as needed
}

# Specify the paths of the applications you want to check along with their display names
$applicationPaths = @{
   "C:\Program Files (x86)\3CXPhone\3CXPhone.exe" = "3CXPhone"
   "C:\Program Files\7-Zip\7zFM.exe" = "7-Zip"
   "C:\Program Files\FileZilla FTP Client\filezilla.exe" = "FileZilla Client"
   "C:\Program Files\dhcp\dhcpctl.exe" = "HaneWin DHCP"   
   "C:\Program Files (x86)\Batch Configuration\Batch Configuration\Batch Configuration.exe" = "HikVision Batch Configuration"
   "C:\Program Files\NPortAdminSuite\bin\npadmer.exe" = "MOXA NPortAdminSuite"
   "C:\Program Files\Moxa\NPortAdminSuite\bin\npadmer.exe" = "MOXA NPortAdminSuite"
   "C:\Program Files\Notepad++\notepad++.exe" = "Notepad++"
   "C:\Program Files\Npcap\NPFInstall.exe" = "Npcap"
   "C:\Program Files (x86)\PRTG Network Monitor\PRTG Probe.exe" = "PRTG Probe"
   "C:\Program Files\PuTTY\putty.exe" = "PuTTY"
   "C:\Program Files\VMware\VMware Tools\vmtoolsd.exe" = "VMWare Tools"
   "C:\Program Files\RealVNC\VNC Server\vncserver.exe" = "VNC Server"
   "C:\Program Files\RealVNC\VNC Viewer\vncviewer.exe" = "VNC Viewer"
   "C:\Program Files (x86)\WinSCP\WinSCP.exe" = "WinSCP"
   "C:\Program Files\Wireshark\Wireshark.exe" = "Wireshark"
	# Add more application paths as needed
}

# Generate a timestamp for the output file
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"

# Combine the custom output directory and the output folder name
$outputDirectory = Join-Path -Path $CustomOutputDirectory -ChildPath "Geinstalleerde Software"
$outputFile = Join-Path -Path $outputDirectory -ChildPath "$timestamp.txt"

# Create the output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory
}

# Initialize an array to store file and application version information
$fileVersions = @()

# Loop through each file path and get file version
foreach ($filePath in $filePaths.Keys) {
    if (Test-Path $filePath) {
        $fileName = $filePaths[$filePath]
        $fileVersion = (Get-Command $filePath).FileVersionInfo.FileVersion
        if ($fileVersion) {
            $fileVersions += "{0,-50} {1}" -f $fileName, $fileVersion
        } else {
            $fileVersions += "$fileName : No version information available"
        }
    }
}

# Add a separator between file and application information
$fileVersions += "---------------------------------------"

# Loop through each application path and get application version
foreach ($appPath in $applicationPaths.Keys) {
    if (Test-Path $appPath) {
        $appName = $applicationPaths[$appPath]
        $appVersion = (Get-Command $appPath).FileVersionInfo.FileVersion
        if ($appVersion) {
            $fileVersions += "{0,-50} {1}" -f $appName, $appVersion
        } else {
            $fileVersions += "$appName : No version information available"
        }
    }
}

# Add a separator between application and Windows information
$fileVersions += "---------------------------------------"

# Get Windows version information
$windowsVersion = Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object Caption, Version, BuildNumber

# Get installed updates and hotfixes
$updates = Get-HotFix | Select-Object Description, HotFixID

# Add Windows version information to the output
$fileVersions += "Windows Version Information:"
$fileVersions += "Caption: $($windowsVersion.Caption)"
$fileVersions += "Version: $($windowsVersion.Version)"
$fileVersions += "Build Number: $($windowsVersion.BuildNumber)"
$fileVersions += ""

# Add installed updates and hotfixes to the output
$fileVersions += "Installed Updates and Hotfixes:"
foreach ($update in $updates) {
    $fileVersions += "$($update.Description) : $($update.HotFixID)"
}
$fileVersions += ""

# Write file, application, and Windows information to the output file
$fileVersions | Out-File -FilePath $outputFile

Write-Host "Information has been written to $outputFile"
