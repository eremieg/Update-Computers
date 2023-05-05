[CmdletBinding()]
param()

# Import required modules
try {
    Import-Module ActiveDirectory -ErrorAction Stop
}
catch {
    Write-Error "Failed to import ActiveDirectory module. Please ensure the module is installed."
    exit 1
}

try {
    Import-Module PSWindowsUpdate -ErrorAction Stop -UseSSL
}
catch {
    Write-Error "Failed to import PSWindowsUpdate module. Please ensure the module is installed."
    exit 1
}

# Wake on LAN function
function Send-WOL {
    param(
        [parameter(Mandatory = $true)][string]$MACAddress
    )

    $broadcast = [Net.IPAddress]::Broadcast
    $mac = $MACAddress -replace "[:-]", ""
    $packet = ([Byte[]](, 0xFF * 6)) + (([Byte[]](, [Convert]::ToByte($mac.Substring(0, 2), 16)) * 16) * 6)

    $udpClient = New-Object Net.Sockets.UdpClient
    $udpClient.Connect($broadcast, 7)
    $udpClient.Send($packet, 102)
    $udpClient.Close()
}

# Prompt user for the CSV file path
$csvPath = Read-Host -Prompt 'Please provide the path to the CSV file'
if (-not (Test-Path $csvPath)) {
    Write-Error "CSV file not found at the provided path."
    exit 1
}

try {
    $computers = Import-Csv $csvPath -ErrorAction Stop
}
catch {
    Write-Error "Failed to import CSV file. Error: $_"
    exit 1
}

if (-not (($computers | Get-Member -MemberType NoteProperty).Name -contains 'ComputerName')) {
    Write-Error "The CSV file is missing the 'ComputerName' column."
    exit 1
}

# Initialize result arrays
$notAwakenedComputers = @()
$updatedComputers = @()
$updateErrorComputers = @()

# Create a log file
$logFile = "UpdateComputers.log"
if (Test-Path $logFile) {
    Remove-Item $logFile
}
New-Item $logFile -ItemType "file"

function Write-Log([string]$Message, [string]$Level = "INFO") {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "[$timestamp] [$Level] $Message"
}

# Get MAC address from Active Directory
function Get-MACAddressFromAD($computerName) {
    try {
        $computer = Get-ADComputer $computerName -Properties *
        $macAddress = $computer.'msDS-PhyShardwareId'
        return $macAddress
    }
    catch {
        Write-Warning "Failed to retrieve MAC address for $computerName from Active Directory. Error: $_"
        return $null
    }
}

# Iterate through each computer in the CSV file
$computers | ForEach-Object -Parallel -ThrottleLimit 4 {
    $computerName = $_.ComputerName
    $macAddress = Get-MACAddressFromAD $computerName
    $retryCount = 0
    $maxRetries = 4
    $computerAwake = $false
    $result = New-Object -TypeName PSObject -Property @{
        ComputerName     = $computerName
        UpdatesInstalled = @()
        Awake            = $false
        UpdateError      = $false
    }

    while (-not $computerAwake -and $retryCount -lt $maxRetries) {
        # Wake up the computer using its MAC address
        if ($macAddress) {
            Write-Verbose "Attempting to wake up $computerName (attempt $(($retryCount + 1)))..."
            Write-Log "Attempting to wake up $computerName (attempt $(($retryCount + 1)))..."
            Send-WOL -MACAddress $macAddress
            Start-Sleep -Seconds 60
        }
        else {
            Write-Warning "Skipping $computerName due to missing MAC address."
            Write-Log "Skipping $computerName due to missing MAC address." -Level "WARNING"
            break
        }

        # Check if the computer is online using Test-Connection
        Write-Verbose "Checking if $computerName is online..."
        if (Test-Connection -ComputerName $computerName -Count 1 -Quiet) {
            $computerAwake = $true
            $result.Awake = $true
        }
        else {
            $retryCount++
        }
    }
    
    if ($computerAwake) {
        # If the computer is online, update it using Invoke-WUInstall
        Write-Verbose "$computerName is online. Updating..."
        
        $updateAttempt = 0
        $maxUpdateAttempts = 3
        $updateSuccess = $false

        while (-not $updateSuccess -and $updateAttempt -lt $maxUpdateAttempts) {
            try {
                $job = Invoke-WUInstall -ComputerName $computerName -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -AsJob -ErrorAction Stop -TimeoutSec 3600 -Verbose
                $updates = Wait-Job -Job $job | Receive-Job
                # Add the computer and its installed updates to the result object
                $result.UpdatesInstalled = $updates
                $updateSuccess = $true
                $updatedComputers += $computerName
            }
            catch {
                Write-Warning "Failed to update $computerName (attempt $(($updateAttempt + 1))). Error: $_"
                Write-Log "Failed to update $computerName (attempt $(($updateAttempt + 1))). Error: $_" -Level "WARNING"
                $updateAttempt++
                $result.UpdateError = $true
                Start-Sleep -Seconds 300
            }
        }
    }

    if (-not $computerAwake) {
        Write-Verbose "$computerName did not wake up after $maxRetries attempts."
        Write-Log "$computerName did not wake up after $maxRetries attempts." -Level "WARNING"
        $notAwakenedComputers += $computerName
    }
    elseif ($result.UpdateError) {
        $updateErrorComputers += $computerName
    }
}

# Generate HTML output
Write-Verbose "Generating HTML output..."

# Updated computers table
$updatedComputersTable = $updatedComputers | 
Select-Object @{Name = 'Computer Name'; Expression = { $_ }} |
ConvertTo-Html -Fragment

# Not awakened computers table
$notAwakenedComputersTable = $notAwakenedComputers | 
Select-Object @{Name = 'Computer Name'; Expression = { $_ }} | 
ConvertTo-Html -Fragment

# Update error computers table
$updateErrorComputersTable = $updateErrorComputers | 
Select-Object @{Name = 'Computer Name'; Expression = { $_ }} | 
ConvertTo-Html -Fragment

# Generate the complete HTML content
$html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Update Report</title>
    <style>
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        margin: 40px;
        color: #333;
        background-color: #f5f5f5;
    }
    h1 {
        font-size: 24px;
        color: #4a4a4a;
        border-bottom: 1px solid #4a4a4a;
        padding-bottom: 10px;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 30px;
    }
    th, td {
        border: 1px solid #dddddd;
        text-align: left;
        padding: 8px;
    }
    th {
        background-color: #689f38;
        color: white;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .summary {
        font-size: 18px;
        font-weight: bold;
        margin-bottom: 20px;
        color: #4a4a4a;
    }
    div {
        margin-bottom: 5px;
    }
</style>

</head>
<body>
    <h1>Update Report</h1>
    <div class="summary">Summary:</div>
    <div>Successfully updated computers: $($updatedComputers.Count)</div>
    <div>Computers that failed to wake up: $($notAwakenedComputers.Count)</div>
    <div>Computers with update errors: $($updateErrorComputers.Count)</div>
    <h1>Successfully Updated Computers</h1>
    <table>
        $updatedComputersTable
    </table>
    <h1>Computers That Failed to Wake Up</h1>
    <table>
        $notAwakenedComputersTable
    </table>
    <h1>Computers With Update Errors</h1>
    <table>
        $updateErrorComputersTable
    </table>
</body>
</html>
"@

# Save the HTML content to a file
$htmlOutputPath = "UpdateReport.html"
Set-Content -Path $htmlOutputPath -Value $html

Write-Verbose "Report saved to $htmlOutputPath"
