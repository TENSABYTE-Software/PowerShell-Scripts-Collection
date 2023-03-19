[CmdletBinding()]
param(
    [Parameter(Mandatory=$false, Position=0)]
    [string]$ExportPath = './'
)

function Get-BatteryStats() {
    $batteryStatsFormatted = @{}
    $battery = Get-CimInstance -ClassName Win32_Battery -Property *
    if (!$battery) {
        return $null
    }
    $batteryStatsFormatted["DesignCapacity"] = $battery.DesignCapacity
    $batteryStatsFormatted["FullChargeCapacity"] = $battery.FullChargeCapacity
    $batteryStatsFormatted["Status"] = $battery.Status

    return $batteryStatsFormatted
}

function Get-Memory() {
    $memory = Get-CimInstance -ClassName Win32_PhysicalMemory -Property * 
    $memoryOutput = @{
        "Capacity" = $null;
        "Speed" = $null;
        "Modules" = 0;
    }
    $memory | %{
        $memoryOutput.Modules++
        $memoryOutput.Speed = $_.Speed.ToString()
        $memoryOutput.Capacity += $_.Capacity
    }
    return $memoryOutput
}



function parseBatteryReport($path = './battery-report.html') {
    function Make-HTMLObj($path) {
        try {
            $html = New-Object -Com 'HTMLFile'
            $unicode = [System.Text.Encoding]::Unicode.GetBytes($path)
            $html.write($unicode)
        } catch {
            Write-Error 'couldnt parse html file'
        }
    }
    function Parse-Item($data, $reg) {
        for (($i = 0); $i -lt $data.length; $i++) {
            $match = $data[$i].innerHTML -match $($reg.ToString())
            if ($match) {
                return $data[$i].innerHTML
            }
        }
    }

    if ($unicode) {
        try {
            
            $dcEl = Parse-Item ($html.getElementsByTagName('tr')) 'DESIGN CAPACITY'
            $fcEl = Parse-Item ($html.getElementsByTagName('tr')) 'FULL CHARGE CAPACITY'

        }
        catch {
            throw 'Error: Document not found'
        }
    }
}

function Get-DeviceDetails() {
    $machineFormatted = @{}
    
    $machine = Get-CimInstance -ClassName Win32_ComputerSystemProduct -Property Vendor, IdentifyingNumber
    $cpu = Get-CimInstance -ClassName Win32_Processor -Property Name, NumberOfCores, NumberOfLogicalProcessors
    $memory = Get-Memory
    $battery = Get-BatteryStats

    $machineFormatted["Manufacturer"] = $machine.Vendor
    $machineFormatted["SerialNumber"] = $machine.IdentifyingNumber
    $machineFormatted["Model"] = $machine.Name

    $machineFormatted["CPUName"] = $cpu.Name
    $machineFormatted["PhysicalCores"] = $cpu.NumberOfCores
    $machineFormatted["LogicalCores"] = $cpu.NumberOfLogicalProcessors

    $machineFormatted["MemoryCapacity"] = $memory.Capacity
    $machineFormatted["MemorySpeed"] = $memory.Speed
    $machineFormatted["MemoryInstalledModules"] = $memory.Modules

    $machineFormatted["BatteryDesignCapacity"] = $battery.DesignCapacity
    $machineFormatted["BatteryFullChargeCapacity"] = $battery.FullChargeCapacity
    $machineFormatted["BatteryHealth"] = "$($($battery.FullChargeCapacity / $battery.DesignCapacity) * 10)%"

    $outputObj = New-Object PSObject -Property $machineFormatted 
    $machineFormatted 
}


Get-DeviceDetails 