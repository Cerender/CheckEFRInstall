<#------------------------------------------------------------------------------
    Jason McClary
    mcclarj@mail.amc.edu
    28 Mar 2016
    
    Description:
    Check computers for Access install
    
    Arguments:
    If blank script runs against local computer
    Multiple computer names can be passed as a list separated by spaces:
        CheckEFRInstall.ps1 computer1 computer2 anotherComputer
    A text file with a list of computer names can also be passed
        CheckEFRInstall.ps1 comp.txt
        
    Tasks:
    - Use a list of computers to check:
        - Access is installed with new package
        - Install completed
        - Access is version 5.0.15.0
    - Create a CSV containg that data
    - Create a new file of computers not checked
    
    Path to logFile file:
        C:\AMC_Install_Logs\Access\AccessClient_Install.log
        \\XX_COMP_XX\c$\AMC_Install_Logs\Access\AccessClient_Install.log
    
    Path to application:
        C:\Program Files (x86)\Access\Repository.v5\Client\Acs.Repository.Client.exe
        \\XX_COMP_XX\c$\Program Files (x86)\Access\Repository.v5\Client\Acs.Repository.Client.exe
        
------------------------------------------------------------------------------#>
#Date/ Time Stamp
$dtStamp = Get-Date -UFormat "%Y%m%d%H%M%S"

# CONSTANTS
set-variable logFilePath -option Constant -value "$\AMC_Install_Logs\Access\AccessClient_Install.log"
set-variable appPath -option Constant -value "$\Program Files (x86)\Access\Repository.v5\Client\Acs.Repository.Client.exe"
set-variable logOutput -option Constant -value "Status_$dtStamp.csv"

## Format arguments from none, list or text file 
IF (!$args){
    $compNames = $env:computername # Get the local computer name
} ELSE {
    $passFile = Test-Path $args

    IF ($passFile -eq $True) {
        $compNames = get-content $args
    } ELSE {
        $compNames = $args
    }
}

# Create header
"Computer Name,Version,Install Date" | Out-file -filepath $logOutput -Encoding utf8

# Loop through all computers
FOREACH ($compName in $compNames) {
    # Clear the variables
    $appVersion = ""
    $logFile = ""
    
    IF(Test-Connection -count 1 -quiet $compName){                         # Check connection to computer
        # Get the system drive - should be C but you never know...
        $driveLetter = Get-WMIObject -class Win32_OperatingSystem -Computername $compName | select-object SystemDrive
        $currAppPath = "\\$compName\$($driveLetter.SystemDrive[0])$appPath"
        $currLogPath = "\\$compName\$($driveLetter.SystemDrive[0])$logFilePath"
    
        IF (Test-Path $currAppPath) {                                      # Check Path to Appplication
            $appVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($currAppPath).FileVersion
        } ELSE {
            $appVersion = "Not Installed"
        }
        
        IF (Test-Path $currLogPath) {                                      # Check Path to Installer Log
            $logFile = get-content $currLogPath
            $logFile = $logFile[-1]
        }ELSE{
            $logFile = "Log File Not Found"
        }
        
        $logString = "$compName,$appVersion,$logFile"
    } ELSE {
        $logString = "$compName,Could not connect"
    }
    
    # Write results for current computer
    $logString | Out-file -filepath $logOutput -append -Encoding utf8
   
}
