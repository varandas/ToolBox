#REQUIRES -Version 4.0
<#
.Synopsis
   
.DESCRIPTION
   
.EXAMPLE
  
.EXAMPLE
  
.EXAMPLE
  
#>
param( 
    
)
# --------------------------------------------------------------------------------------------
#region HEADER
$SCRIPT_TITLE = "Get-MissingUpdates"
$SCRIPT_VERSION = "1.0"

$ErrorActionPreference 	= "Continue"	# SilentlyContinue / Stop / Continue

# -Script Name: LOG_NAME.ps1------------------------------------------------------ 
# Based on PS Template Script Version: 1.0
# Author: Jose Varandas

#
# Owned By: Jose Varandas
# Purpose: 
#
# Dependencies: 
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
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "EXAMPLE:" -iTabs 1        
    Write-Log -sMessage "============================================================================================================================" -iTabs 1                		
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
		yyy/mm/dd - Jose Varandas - 1.0			
           CHANGELOG:
               -> Script Created
#>							
# -------------------------------------------------------------------------------------------- 
#endregion
# --------------------------------------------------------------------------------------------
#region Standard FUNCTIONS
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
    $global:original = Get-Location
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
    Set-Location $global:original
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
# Standard Variables
    # *****  Change Logging Path and File Name Here  *****    
    $sOutFileName	= "LogNAME.log" # Log File Name    
    $sEventSource   = "ToolBox" # Event Source Name
    # ****************************************************
    $sScriptName 	= $MyInvocation.MyCommand
    $sScriptPath 	= Split-Path -Parent $MyInvocation.MyCommand.Path
    $sLogRoot		= "C:\Logs\System\SCCM"
    $sOutFilePath   = $sLogRoot
    $sLogFile		= Join-Path -Path $SLogRoot -ChildPath $sOutFileName
    $global:iExitCode = 0
    $sUserName		= $env:username
    $sUserDomain	= $env:userdomain
    $sMachineName	= $env:computername
    $sCMDArgs		= $MyInvocation.Line    
    $iLogFileSize 	= 1048576
    # ****************************************************
# Specific Variables
    
    # ****************************************************  
#endregion 
# --------------------------------------------------------------------------------------------
#region MAIN_SUB

Function MainSub{
# ===============================================================================================================================================================================
#region 1_PRE-CHECKS            
    Write-Log -iTabs 1 "Starting 1 - Pre-Checks."-scolor Cyan
    #region 1.0 
    #endregion        
    Write-Log -iTabs 1 "Completed 1 - Pre-Checks."-sColor Cyan    
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================

# ===============================================================================================================================================================================
#region 2_EXECUTION
    Write-Log -iTabs 1 "Starting 2 - Execution." -sColor cyan    
    #region 2.1        
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

# Quiting with exit code
Exit $global:iExitCode
#endregion