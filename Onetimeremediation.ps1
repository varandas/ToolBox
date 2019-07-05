# REQUIRES -Version 4.0
param(
    [ValidateSet("Check","Run","Auto-Run")][string]$Action="Check"
)
# --------------------------------------------------------------------------------------------
#region HEADER
$SCRIPT_TITLE = "OneTimeRemediationDeliver"
$SCRIPT_VERSION = "1.0"

$ErrorActionPreference 	= "Continue"	# SilentlyContinue / Stop / Continue

# -Script Name: OneTimeRemediationDeliver.ps1------------------------------------------------------ 
# Based on PS Template Script Version: 1.0
# Author: Jose Varandas

#
# Owned By: Jose Varandas
# Purpose: Use P4AOFC01 account to deliver custom Scheduled Tasks for remediation.
#
# Dependencies: 
#                ID running script must be P4AOFC01
#
# Known Issues: 
#
# Arguments: 
Function How-ToScript(){
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "NAME:" -iTabs 1
        Write-Log -sMessage ".\$sScriptName " -iTabs 2     
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "ARGUMENTS:" -iTabs 1            
            Write-Log -sMessage "-Action (Check/Run/Auto-Run) -> Defines Script Execution Mode" -iTabs 3        
                Write-Log -sMessage "-> Check (Default)-> Script will run Pre-checks and Pos-Checks. No Exceution" -iTabs 4        
                Write-Log -sMessage "-> Run -> Runs script (Pre-Checks,Excecution,Post-Checks)" -iTabs 4
                Write-Log -sMessage "-> Auto-Run -> Runs script accepting default options while running Pre-Checks,Excecution,Post-Checks" -iTabs 4  
            Write-Log -sMessage "-PCList ->  Patch to target PC List" -iTabs 3        
                Write-Log -sMessage "TXT File, one PC per line" -iTabs 4                        
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "EXAMPLE:" -iTabs 1
        Write-Log -sMessage ".\$sScriptName -Action Check" -iTabs 2     
            Write-Log -sMessage "Script will run all Pre-Checks. No Changes will happens PC List empty makes script run action locally" -iTabs 2     
        Write-Log -sMessage ".\$sScriptName -Action Run -PCList `"C:\XOM\SPTTemp\pclist.txt`"" -iTabs 2     
            Write-Log -sMessage "Script will run all steps remediations in PCs listed" -iTabs 2  
        Write-Log -sMessage ".\$sScriptName -Action Auto-Run -PCList `"C:\XOM\SPTTemp\pclist.txt`"" -iTabs 2     
            Write-Log -sMessage "Script will run all steps remediations in PCs listed. If any confirmation is needed, obvious answers will be assumed." -iTabs 2        
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
#		
}
#endregion
#region EXIT_CODES
<# Exit Codes:
            0 - Script completed successfully

            3xxx - SUCCESS

            5xxx - INFORMATION          

            7xxx - WARNING            

            9XXX - ERROR
           
            9999 - Unhandled Exception     

   
 Revision History: (Date, Author, Version, Changelog)
		2019/07/05 - Jose Varandas - 1.0			
           CHANGELOG:
               -> Script Created
#>							
# -------------------------------------------------------------------------------------------- 
#endregion
# --------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------
#region FUNCTIONS
Function Start-Log(){	
# --------------------------------------------------------------------------------------------
# Function StartLog

# Purpose: Checks to see if a log file exists and if not, created it
#          Also checks log file size
# Parameters:
# Returns: None
# --------------------------------------------------------------------------------------------
    #Check to see if the log folder exists. If not, create it.
    If (!(Test-Path $sOutFilePath )) {
        New-Item -type directory -path $sOutFilePath | Out-Null
    }
    #Check to see if the log file exists. If not, create it
    If (!(Test-Path $sLogFile )) {
        New-Item $sOutFilePath -name $sOutFileName -type file | Out-Null
    }
	Else
	{
        #File exists, check file size
		$sLogFile = Get-Item $sLogFile
        
        # Check to see if the file is > 1 MB and purge if possible
        If ($sLogFile.Length -gt $iLogFileSize) {
            $sHeader = "`nMax file size reached. Log file deleted at $global:dtNow."
            Remove-Item $sLogFile  #Remove the existing log file
            New-Item $sOutFilePath -name $sOutFileName -type file  #Create the new log file
        }
    }    
    Write-Log $sHeader -iTabs 0  
	Write-Log -sMessage "############################################################" -iTabs 0 
    Write-Log -sMessage "" -iTabs 0 
    Write-Log -sMessage "============================================================" -iTabs 0 	
    Write-Log -sMessage "$SCRIPT_TITLE ($sScriptName) $SCRIPT_VERSION - Start" -iTabs 0 -bEventLog $true -iEventID 5003 -sSource $sEventSource
	Write-Log -sMessage "============================================================" -iTabs 0 
	Write-Log -sMessage "Script Started at $(Get-Date)" -iTabs 0 
	Write-Log -sMessage "" -iTabs 0     
	Write-Log -sMessage "Variables:" -iTabs 0 
	Write-Log -sMessage "Script Title.....:$SCRIPT_TITLE" -iTabs 1 
	Write-Log -sMessage "Script Name......:$sScriptName" -iTabs 1 
	Write-Log -sMessage "Script Version...:$SCRIPT_VERSION" -iTabs 1 
	Write-Log -sMessage "Script Path......:$sScriptPath" -iTabs 1
	Write-Log -sMessage "User Name........:$sUserDomain\$sUserName" -iTabs 1
	Write-Log -sMessage "Machine Name.....:$sMachineName" -iTabs 1
	Write-Log -sMessage "Log File.........:$sLogFile" -iTabs 1
	Write-Log -sMessage "Command Line.....:$sCMDArgs" -iTabs 1
    Write-Log -sMessage "Arguments===================================================" -iTabs 0 
	Write-Log -sMessage "-DebugLog...:$DebugLog" -iTabs 1
    Write-Log -sMessage "-NoRelaunch.:$NoRelaunch" -iTabs 1 
    Write-Log -sMessage "-Action.....:$Action" -iTabs 1 
    Write-Log -sMessage "-Scope:$Scope" -iTabs 1    
	Write-Log -sMessage "============================================================" -iTabs 0    
}           ##End of Start-Log function
Function Write-Log(){
# --------------------------------------------------------------------------------------------
# Function Write-Log

# Purpose: Writes specified text to the log file
# Parameters: 
#    sMessage - Message to write to the log file
#    iTabs - Number of tabs to indent text
#    sFileName - name of the log file (optional. If not provied will default to the $sLogFile in the script
# Returns: None
# --------------------------------------------------------------------------------------------
    param( 
        [string]$sMessage="", 
        [int]$iTabs=0, 
        [string]$sFileName=$sLogFile,
        [boolean]$bTxtLog=$true,
        [boolean]$bConsole=$true,
        [string]$sColor="white",         
        [boolean]$bEventLog=$false,        
        [int]$iEventID=0,
        [ValidateSet("Error","Information","Warning")][string]$sEventLogType="Information",
        [string]$sSource=$sEventIDSource        
    )
    
    #Loop through tabs provided to see if text should be indented within file
    $sTabs = ""
    For ($a = 1; $a -le $iTabs; $a++) {
        $sTabs = $sTabs + "    "
    }

    #Populated content with timeanddate, tabs and message
    $sContent = "||"+$(Get-Date -UFormat %Y-%m-%d_%H:%M:%S)+"|"+$sTabs + "|"+$sMessage

    #Write content to the file
    if ($bTxtLog){
        Add-Content $sFileName -value  $sContent -ErrorAction SilentlyContinue
    }    
    #write content to Event Viewer
    if($bEventLog){
        try{
            New-EventLog -LogName Application -Source $sSource -ErrorAction SilentlyContinue
            if ($iEventID -gt 9000){
                $sEventLogType = "Error"
            }
            elseif ($iEventID -gt 7000){
                $sEventLogType = "Warning"
            }
            else{
                $sEventLogType = "Information"
            }
            Write-EventLog -LogName Application -Source $sSource -EntryType $sEventLogType -EventId $iEventID -Message $sMessage -ErrorAction SilentlyContinue
        }
        catch{
            
        }
    }
    # Write Content to Console
    if($bConsole){        
            Write-Host $sContent -ForegroundColor $scolor        
    }
	
}           ##End of Write-Log function
Function Finish-Log(){
# --------------------------------------------------------------------------------------------
# Function EndLog
# Purpose: Writes the last log information to the log file
# Parameters: None
# Returns: None
# --------------------------------------------------------------------------------------------
    #Loop through tabs provided to see if text should be indented within file
	Write-Log -sMessage "" -iTabs 0 
    Write-Log -sMessage "$SCRIPT_TITLE ($sScriptName) $SCRIPT_VERSION Completed at $(Get-date) with Exit Code $global:iExitCode - Finish" -iTabs 0  -bEventLog $true -sSource $sEventSource -iEventID $global:iExitCode    
    Write-Log -sMessage "============================================================" -iTabs 0     
    Write-Log -sMessage "" -iTabs 0     
    Write-Log -sMessage "" -iTabs 0 
    Write-Log -sMessage "" -iTabs 0 
    Write-Log -sMessage "" -iTabs 0     
}             ##End of End-Log function
function ConvertTo-Array{
    begin{
        $output = @(); 
    }
    process{
        $output += $_;   
    }
    end{
        return ,$output;   
    }
}
#endregion
# --------------------------------------------------------------------------------------------
#region SPECIALTY-FUNCTIONS
Function Test-Conectivity {
     Param(
        [string]$comp="localhost"
    )
    $ping = Test-Connection -ComputerName $comp -Count 1 -Quiet
    if ($ping){
        return $true
    }
    else{
        return $false
    }       
}
Function Manage-RemoteRegistry{
    Param(
            [string]$comp="localhost",
            [ValidateSet("Start","Stop")][string]$action="Stop"
        )
    if ($action -eq "Start" ){
        #capture initial Svc Start Mode
        $global:initialRemoteRegSvcStartMode = (Get-WmiObject -Class Win32_Service -Property StartMode -Filter "Name='RemoteRegistry'" -ComputerName CTC6L33K91642).StartMode
        #If Svc is disabled, enable it as manual
        If ($global:initialRemoteRegSvcStartMode -eq "Disabled"){
            Set-Service  -ComputerName $comp -Name RemoteRegistry -StartupType Manual
        }
        #Check Svc Status
        $svc = (Get-Service -ComputerName $comp -Name RemoteRegistry).Status

        #If not running, start to run it
        if ($svc -eq "Stopped"){
            Get-Service -ComputerName $comp -Name RemoteRegistry | Start-Service 
        }
    }
    else{
        $svc = (Get-Service -ComputerName $comp -Name RemoteRegistry).Status
        if ($svc -eq "Running"){
            Get-Service -ComputerName $comp -Name RemoteRegistry | Stop-Service -Force
        }
        if ($global:initialRemoteRegSvcStartMode -ne $null){
            $ifinalRemoteRegSvcStartMode = (Get-WmiObject -Class Win32_Service -Property StartMode -Filter "Name='RemoteRegistry'" -ComputerName CTC6L33K91642).StartMode
            If ($global:initialRemoteRegSvcStartMode -ne $ifinalRemoteRegSvcStartMode){
                Set-Service  -ComputerName $comp -Name RemoteRegistry -StartupType $global:initialRemoteRegSvcStartMode
            }
        }
    }
}
function Execute-OneTimeRemediation {
Param(
            [string]$comp="localhost",
            [string]$source="\\HOURDS750\TEMP\Dummy",
            [string]$destination="\\$comp\c$\xom\SPTTemp\CVE2019-0708"
        )
#Copy PSScript into SPTTemp
    robocopy.exe $source $destination /MIR /r:10 /w:5

    #Create Scheduled Task with PS Script
    schtasks /create /tn OneTimeFix /tr "c:\xom\SPTTemp\CVE2019-0708\CVE2019-0708-Remediation.ps1" /ru System /SC ONCE /ST 00:01 /s $computername

    #Execute Scheduled Task
    schtasks /run /tn OneTimeFix /s \\$computername

    #Delete Scheduled Task
    schtasks /delete /tn OneTimeFix /s \\$computername
}
#endregion

#region VARIABLES
# Standard Variables
    # *****  Change Logging Path and File Name Here  *****    
    $sOutFileName	= "OneTimeRemediationDeliver.log" # Log File Name    
    $sEventSource   = "DWS_Script" # Event Source Name
    # ****************************************************
    $sScriptName 	= $MyInvocation.MyCommand
    $sScriptPath 	= Split-Path -Parent $MyInvocation.MyCommand.Path
    $sLogRoot		= "C:\XOM\Logs\System\OneTimeRemediationDeliver\"    
    $sOutFilePath   = $sLogRoot
    $sLogFile		= Join-Path -Path $SLogRoot -ChildPath $sOutFileName
    $global:iExitCode = 0
    $sUserName		= $env:username
    $sUserDomain	= $env:userdomain
    $sMachineName	= $env:computername
    $sCMDArgs		= $MyInvocation.Line
    $bAllow64bitRelaunch = $true
    $iLogFileSize 	= 1048576
    # ****************************************************
    # Script Specific Variables
    $global:initialRemoteRegSvcStartMode=$null
        
#endregion 
# --------------------------------------------------------------------------------------------
#region MAIN_SUB

Function MainSub{
    $computers = @('CTC6L33KRF542','CTC6L33K91642','CTC6L47C6NC72','CTC6L47C7JC72','CTC6TA062379171')

    Write-Log -iTabs 1 "Loading Device List..."
    Write-Log -iTabs 2 $Computers
    foreach ($computername in $computers){
        Write-Log -itabs 1 "Starting to work on $ComputerName"
        $output = 0
        #Test-Connection
            Write-Log -itabs 2 "Checking Connectivity"
            $canReach = Test-Conectivity -comp $computername     
            if ($canReach){
                Write-Log -itabs 3 "Device found online" -sColor Green
            }
            else{
                Write-Log -itabs 3 "Unable to reach Device. Skipping" -sColor Red
                continue            
            }
        #Enable RemoteRegistry
            Write-Log -itabs 2 "Enabling RemoteRegistry"
            try{
                Manage-RemoteRegistry $computerName -Action Start
                Write-Log -iTabs 3 "RemoteRegistry enabled" -sColor Green
            }
            catch{
                Write-Log -itabs 3 "Error enabling RemoteRegistry" -sColor Red
                continue            
            }
    
        #Run Payload
            #Execute-OneTimeRemediation $computername \\HOURDS750\TEMP\CVE2019-0708 \\$computername\c$\xom\SPTTemp\CVE2019-0708
    
        #Disable RemoteRegistry    
           Write-Log -itabs 2 "Disabling RemoteRegistry"
            try{
                Manage-RemoteRegistry $computerName -Action Stop
                Write-Log -iTabs 3 "RemoteRegistry disabled" -sColor Green
            }
            catch{
                Write-Log -itabs 3 "Error enabling RemoteRegistry" -sColor Red
                continue            
            }
    }
}


#endregion
# --------------------------------------------------------------------------------------------

# --------------------------------------------------------------------------------------------
#region MAIN_PROCESSING

# Starting log
$global:original = Get-Location  
Start-Log
Try {
	MainSub    
}
Catch {
	# Log a general exception error
	Write-Log -sMessage "Error running script" -iTabs 0        
    if ($global:iExitCode -eq 0){
	    $global:iExitCode = 9999
    }                
}
# Stopping the log
Finish-Log
Set-Location $global:original
# Quiting with exit code
Exit $global:iExitCode
#endregion