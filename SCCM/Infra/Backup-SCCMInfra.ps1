
##########################################################################################################################################

#REQUIRES -Version 4.0
<#
.Synopsis
   Script responsible for keeping permafrost information always updated
.DESCRIPTION
   Copy CAS and AM1 backup to Permafrost. Copy CAS-SCCMContentLib into permafrost. Copy Daldat01 source locations into permafrost
.EXAMPLE
  .\Backup-SCCMInfra.ps1
    Infers "-Action Continue"
        Script will check local server registry for pending copy actions and resume backing up information into permafrost
.EXAMPLE
  .\Backup-SCCMInfra.ps1 -Action "New"    
        Script will consider any permafrost information old and will start copies from scratch.
  
#>
param( 
    
)
# --------------------------------------------------------------------------------------------
#region HEADER
$SCRIPT_TITLE = "Backup-SCCMInfra"
$SCRIPT_VERSION = "2.0"

$ErrorActionPreference 	= "Continue"	# SilentlyContinue / Stop / Continue

# -Script Name: Backup-SCCMInfra.ps1------------------------------------------------------ 
# Based on PS Template Script Version: 1.0
# Author: Gilberto Hepp

#
# Owned By: Jose Varandas
# Purpose: Script responsible for keeping permafrost information always updated
# 
# Dependencies: Account running this script must have read accessin local host and read/write access in permafrost
#
# Known Issues: 
#



Function Show-ScriptUsage(){
# --------------------------------------------------------------------------------------------
# Function Show-ScriptUsage

# Purpose: Show how to use this script
# Parameters: None
# Returns: None
# --------------------------------------------------------------------------------------------
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "NAME:" -iTabs 1
        Write-Log -sMessage ".\$sScriptName " -iTabs 2     
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "ARGUMENTS:" -iTabs 1                        
    Write-Log -sMessage     "-Action:" -iTabs 2        
    Write-Log -sMessage         "New: Considers any info found in Permafrost outdated and starts copy from scratch." -iTabs 3        
    Write-Log -sMessage         "Continue: Reads localhost registry for ongoing copy action and resumes from last stage." -iTabs 3        
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "EXAMPLE:" -iTabs 1    
    Write-Log -sMessage     ".\Backup-SCCMInfra.ps1" -iTabs 2
    Write-Log -sMessage         "Infers `"-Action Continue`" "  -iTabs 3
    Write-Log -sMessage         "Script will check local server registry for pending copy actions and resume backing up information into permafrost" -iTabs 3
    Write-Log -sMessage     ".\Backup-SCCMInfra.ps1 -Action New" -iTabs 2
    Write-Log -sMessage         "Script will consider any permafrost information old and will start copies from scratch.    "  -iTabs 3
    Write-Log -sMessage "============================================================================================================================" -iTabs 1                		
}
#endregion
#region EXIT_CODES
<# Exit Codes:
            0 - Script completed successfully

            3xxx - SUCCESS
            3001 - Script Completed without Errors
            3002 - Script Completed with Continues

            5xxx - INFORMATION     
            5001 - Script start   
            5002 - Script finish

            7xxx - WARNING

            9XXX - ERROR
            
            9999 - Unhandled Exception     

   
 Revision History: (Date, Author, Version, Changelog)
     v1.0 - 03/08/2018 - Gilberto J Hepp
          -> Creation 
     v1.3 - 07/22/2019 - Maria L Ceschin
          -> Added SCCMApplications share to PERMAFROST copies.
     v2.0 - 07/24/2019 - Jose Varandas
          -> Rebranded script to Powershell guidelines
          -> Added Powershell standard header
          -> Added ability to log in Event Viewer
          -> Allow parallel/distributed copying
          -> Allow Resume Copy
          -> Store progress in Server registry (resilience to reboots)
		
#>							
# -------------------------------------------------------------------------------------------- 
#endregion
# --------------------------------------------------------------------------------------------
#region Standard FUNCTIONS
Function Start-Log(){	
# --------------------------------------------------------------------------------------------
# Function Start-Log

# Purpose: Checks to see if a log file exists and if not, created it. Also checks log file size
# Parameters: None
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
        
        # Check to see if the file is > $iLogFileSize and purge if possible
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
    Write-Log -sMessage "$SCRIPT_TITLE ($sScriptName) $SCRIPT_VERSION - Start" -iTabs 0 -bEventLog $true -iEventID 5001 -sSource $sEventSource
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
    Write-Log -sMessage "-Action.....:$Action" -iTabs 1     
	Write-Log -sMessage "============================================================" -iTabs 0    
}           ##End of Start-Log function
Function Write-Log(){
# --------------------------------------------------------------------------------------------
# Function Write-Log

# Purpose: Writes specified text to the log file
# Returns: None
# --------------------------------------------------------------------------------------------
    param( 
        [string]$sMessage="", # Message to be written in Log
        [int]$iTabs=0,        # Tabs before starting to write $sMessage
        [string]$sFileName=$sLogFile, #Log Full Path
        [boolean]$bTxtLog=$true, #Write info to Log        

        [boolean]$bEventLog=$false, #write into to Event Viewer        
        [int]$iEventID=0,           #Event ID
        [ValidateSet("Error","Information","Warning")][string]$sEventLogType="Information", #Event Type
        [string]$sSource=$sEventSource,     #event Source   

        [boolean]$bConsole=$true,#Write info to Console
        [string]$sColor="white" #Info color (Console only)        
        
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
Function Stop-Log(){
# --------------------------------------------------------------------------------------------
# Function Stop-Log
# Purpose: Writes the last log information to the log file
# Parameters: None
# Returns: None
# --------------------------------------------------------------------------------------------
    #Loop through tabs provided to see if text should be indented within file
	Write-Log -sMessage "" -iTabs 0 
    Write-Log -sMessage "$SCRIPT_TITLE ($sScriptName) $SCRIPT_VERSION Completed at $(Get-date) with Exit Code $global:iExitCode - Finish" -iTabs 0  -bEventLog $true -sSource $sEventSource -iEventID 5002  
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
#region Specific FUNCTIONS

#endregion
# --------------------------------------------------------------------------------------------
#region VARIABLES
# Common Variables
    # *****  Change Logging Path and File Name Here  *****    
    $sOutFileName        = "RRRe-LocalLog.log" # Log File Name        
    $sLogRoot		     = "C:\RRRe-Backup" #Log Path Location
    $sEventSource        = "DWS-D3S-RRRe" # Event Source Name
    # ****************************************************
    $global:iExitCode = 0
    $sScriptName 	= $MyInvocation.MyCommand
    $sScriptPath 	= Split-Path -Parent $MyInvocation.MyCommand.Path    
    $sOutFilePath   = $sLogRoot
    $sLogFile		= Join-Path -Path $sLogRoot -ChildPath $sOutFileName    
    $sUserName		= $env:username
    $sUserDomain	= $env:userdomain
    $sMachineName	= $env:computername
    $sCMDArgs		= $MyInvocation.Line    
    $iLogFileSize 	= 10485760
    # ****************************************************
# Specific Variables
    $paths = @(                
        <#
        Scope -> Data owner. Values: SCCMInfra, OSD, Packaging, SecurityUpdates
        #>
        [pscustomobject]@{"Scope"="SCCMInfra";"Host"="DALDAT01";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""},
        [pscustomobject]@{"Scope"="";"Host"="";"SourcePath"="";"DestinationPath"=""}
    )
    $scriptTimeThreshold = 12 # How many hours this script should run before generating an alert
    $regPath = "HKEY_LOCAL_MACHINE\SOFTWARE\ExxonMobil\RRRe"
    # ****************************************************  
#endregion 
# --------------------------------------------------------------------------------------------
#region MAIN_SUB

Function MainSub{
# ===============================================================================================================================================================================
#region 1_PRE-CHECKS            
    Write-Log -iTabs 1 "Starting 1 - Pre-Checks."-scolor Cyan
    #region 1.1 Test LocalHost perms and remote destination perms
    #endregion        
    #region 1.2 Test remote log ability on CAS
    #endregion        
    #region 1.3 Test/Create Registry info on local server
    #endregion        
    #region 1.4 Get last copy status
    #endregion 
    #region 1.5 Set copy stage
        #if new Treat local logs
            #if new+error -> Event to Engineering group and proceed
            #if new+ok/null -> Start copy normally from begining
        #if continue , test local log lcoation
            #if continue+error -> Go to pending step
            #if continue+copy complete -> Terminate Script
    #endregion 
    Write-Log -iTabs 1 "Completed 1 - Pre-Checks."-sColor Cyan    
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================

# ===============================================================================================================================================================================
#region 2_EXECUTION
    Write-Log -iTabs 1 "Starting 2 - Execution." -sColor cyan    
    #region 2.1 If SiteServer, send last SCCM Backup to Permafrost           
        #get status from latest backup
        #2.1.1 move into templocation
        #2.1.2 clean-up
            #remove Scorch logs
            #compress SCCM Infra Logs
            #remove crashdump logs
        #get size
        #2.1.4 send to permafrost
            #if error, set event error and reg error
    #endregion  
    #region 2.2 - If CAS, Copy SCCMContentLib, SMSPKG, SMSPKGSIG to permafrost
        #2.2.1 Get ContentLib Size
        #copy to permafrost 
            #if error, set event error and reg error
        #2.2.2 Get SMSPKG Size
        #copy to permafrost 
            #if error, set event error and reg error
        #2.2.3 Get SMSPKGSIG Size
        #copy to permafrost 
            #if error, set event error and reg error
    #endregion      
    #region 2.3 - If DALDAT01, Copy SCCMPAckages, SCCMApplications to permafrost
        #2.3.1 Get SCCMPAckages Size
        #copy to permafrost 
            #if error, set event error and reg error
        #2.3.2 Get SCCMApplications Size
        #copy to permafrost 
            #if error, set event error and reg error        
    #endregion     
    #region 2.4 - If DACDAT01, Copy OSD ISO paths  to permafrost
        #2.4.X Get OSD ISO Size
        #copy to permafrost 
            #if error, set event error and reg error           
    #endregion     
    Write-Log -iTabs 1 "Completed 2 - Execution." -sColor cyan
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================
        
# ===============================================================================================================================================================================
#region 3_POST-CHECKS
# ===============================================================================================================================================================================
    Write-Log -iTabs 1 "Starting 3 - Post-Checks."-sColor cyan
    #region 3.1     
        #List all Sizes of copies
            #if this is not the first run, compare size growth and time 
        #Create Eventlog for Success
    #endregion 
    Write-Log -iTabs 1 "Completed 3 - Post-Checks."-sColor cyan
    Write-Log -iTabs 0 "" -bConsole $true
#endregion
# ===============================================================================================================================================================================

} #End of MainSub

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
Stop-Log
Set-Location $global:original

# Quiting with exit code
Exit $global:iExitCode
#endregion