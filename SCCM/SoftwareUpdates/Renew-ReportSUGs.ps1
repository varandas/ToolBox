#REQUIRES -Version 4.0
<#
.Synopsis
   Renew-ReportSUGs
.DESCRIPTION
   Script is meant to Renew Reporting SUGs with latest information. It is built to ensure reporting is as accurate as possible from the defined criterias of severity, update classification and product.   
   
   Starting concepts:
    -> MSFT SUG -> Contains ALL valid updates from Microsoft respecting criterias defined as:
        Release Date
        Severity
        Product
        Update Classification
    -> Blacklist SUG -> Contains all updates banned from environment. These will be ignored in all script logic
    -> Report SUG -> Contais all updates from MSFT SUG, which are Deployed and are not part of Blacklist SUG
    -> Missing SUG -> Contains all updates from MSFT SUG not found in Report SUG or Blacklist SUG. The goal is to have Missing SUG always zero

   INPUT: Is managed in the "Script Specific variables" under the "$Scope" switch. If desired, info can be added to switch block and "$Scope" parameter.
    Such action will script to run with "Action" Auto-Run in a more automated manner.
        
        => Information to be added in Switch block. If any is missing, Administrator will be required to add during script execution.
        -> SMS Provider Server Name: Server running SMSProvider. Usually, but not always, is the Central Site Server or the Primary Site Server
        -> SCCM SiteCode: Site code in which Actions are targeted to
        -> TemplateName: <TemplateName> will act as a filter to target only desired SUGS/Deployment Packages/ADRs
            e.g.: TemplateName = "Server-" 
                Script will look for Deployment packages "Server-Montlhy" and "Server-Aged". If they are not found, they will be created.
                Script will look for SUGS names "Server-ADR YYYY-MM-DD HH:MM" These will be considered "Monthly SUGs"
                Script will expect SUG Server-Aged. If it is not found, it will create it.
        Release Date tracking -> Age of an update to be added into Reporting SUGs
        Severity -> Minimum Severity for updates to be tracked
        Product -> Products included in the logic
        Update Classification -> Classifications included in the logic
            
    All actions will be recorded in SCCM Server Logs folder under the name "Renew-ReportSUGs.log"

.EXAMPLE
   .\Renew-ReportSUGs.ps1
        -> IMPLIED PARAMETER: -Scope Other    -> Script will ask additional information in order to execute
        -> IMPLIED PARAMETER: -Action Check   -> Script will not take maintenance actions. Some action might be required (create SUGs or Create Deployment Packages).        
.EXAMPLE
   .\Renew-ReportSUGs.ps1 -Scope MySCCM -Action Run
        -> PARAMETER: -Scope MySCCM   -> Script will use information listed under "MySCCM" variable block. Additional information might be required in order 
            to execute
        -> PARAMETER: -Action Run     -> Script will take maintenance actions, deleting/downloading updates, moving updates between SUGs, deleting SUGs.                 
.EXAMPLE
   .\Optimize-DeployedSUGs.ps1 -Scope MySCCM -Action Auto-Run
        -> PARAMETER: -Scope MySCCM   -> Script will us information listed under "MySCCM" variable block.
        -> PARAMETER: -Action Auto-Run-> Script will take maintenance actions, deleting/downloading updates, moving updates between SUGs, deleting SUGs. 
            No additional information will be asked from Admin. Missing information will cause the script to abort.        
#>
param( 
    [ValidateSet("Check","Run","Auto-Run")][string]$Action="Check",
    [ValidateSet("CAS","VAR","PVA","Other")][string]$Scope="Other"
)
# --------------------------------------------------------------------------------------------
#region HEADER
$SCRIPT_TITLE = "Renew-ReportSUGs"
$SCRIPT_VERSION = "1.0"

$ErrorActionPreference 	= "Continue"	# SilentlyContinue / Stop / Continue

# -Script Name: Renew-ReportSUGs.ps1------------------------------------------------------ 
# Based on PS Template Script Version: 1.0
# Author: Jose Varandas
#
# Owned By: Jose Varandas
# Purpose: Ensure Report SUGs contain accurate Updates.
#
#
# Dependencies: 
#                ID running script must be SCCM administrator
#                SCCM Powershell Module
#                ID running script must be able to reach SMSPRoviderWMI
#                Script must run locally in SMSProvider Server
#                SCCM Current Branch 1802 or higher
#
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
            Write-Log -sMessage "-Scope (CAS/VAR/PVA/Other) -> Defines which SCCM Scope will be targeted." -iTabs 3        
                Write-Log -sMessage "-> CAS: Script targets SCCM01.zlab.varandas.com as Central Server/WMIProvider and WKS-SecurityUpdates as naming convention" -iTabs 4                        
                Write-Log -sMessage "-> VAR: Script targets SCCM01.vlab.varandas.com as Central Server/WMIProvider and VAR as naming convention" -iTabs 4
                Write-Log -sMessage "-> PVA: Script targets SCCM01.plab.varandas.com as Central Server/WMIProvider and VAR as naming convention" -iTabs 4
                Write-Log -sMessage "-> Other: Script doesn't have a target server or naming convention and will require all info to be entered manually." -iTabs 4                   
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "EXAMPLE:" -iTabs 1
        Write-Log -sMessage ".\$sScriptName -Action Check -Scope CAS" -iTabs 2     
            Write-Log -sMessage "Script will run all Pre-Checks in CAS environment. No Changes will happen to the device. Action Argument will not be used with `"-Behavior Check`"" -iTabs 2     
        Write-Log -sMessage ".\$sScriptName -Action Run -Scope VAR" -iTabs 2     
            Write-Log -sMessage "Script will run all coded remediations, pre and  post checks in VAR Environment." -iTabs 2  
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "NOTE:" -iTabs 1
        Write-Log -sMessage "Action Auto-Run is not supported with Scope Other" -iTabs 2                 
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
#		
}
#endregion
#region EXIT_CODES
<# Exit Codes:
            0 - Script completed successfully

            3xxx - SUCCESS

            5xxx - INFORMATION
            5001 - User aborted script
            5002 - User aborted at Aged SUG Creation
            5003 - Script Started

            7xxx - WARNING
            7001 - Error while deleting empty SUG

            9XXX - ERROR
            9001 - Unable to load SCCM PS Module
            9002 - Wrong parameter usage
            9003 - Unable to access SCCM PS Location
            9004 - Unable to query ADRs in SCCM Environment
            9005 - Unable to query SUGs in SCCM Environment
            9006 - Error while creating SUG Aged
            9007 - Error while getting deployment packages
            9008 - Error while creating Deployment packages
            9009 - Error while querying updates from SCCM
            9010 - Error while revieing Aged SUG
            9011 - Error while processing new SUG
            9012 - Error while processing stable SUG
            9013 - Error while processing aged SUG
            9999 - Unhandled Exception     

   
 Revision History: (Date, Author, Version, Changelog)
		2019/04/07 - Jose Varandas - 1.0			
           CHANGELOG:
               -> Script Created
#>							
# -------------------------------------------------------------------------------------------- 
#endregion
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
#region SUP-FUNCTIONS
Function Test-SccmUpdateExpired{
    param(
        [Parameter(Mandatory = $true)]
        $UpdateId,
        $ExpUpdates
    )
    
    # If the WMI query returns more than 0 instances (should NEVER be more than 1 at most), then the update is expired.
    if ($ExpUpdates -match $Updateid){            
        return $true
    }
    else{          
        return $false        
    }
}
Function Test-SccmUpdateSuperseded{
    param(
        [Parameter(Mandatory = $true)]
        $UpdateId,
        $SupUpdate
    )            
    If ($SupUpdate -match $UpdateId){            
        return $true
    }
    else{
        return $false
    }
}
Function Test-SCCMUpdateAge{
    param(
        [Parameter(Mandatory = $true)]
        $UpdateId,            
        $OldUpdates
    )
        
    If ($OldUpdates -match $UpdateId){            
        return $true
    }
    else{
        return $false
    }
}
Function Set-Sug{
# This script is designed to ensure consistent membership of the reporting software update group.
# In this version it is assumed there is only one reporting software update group.  A reporting software
# update group is assumed to never be deployed.  Accordingly, The script will first check to see if the 
# reporting software update group is deployed.  If so the script will display an error and exit.
# If no error then the updates in every other software update group will be reviewed and added to the
# reporting software update group.  There is no check to see if the update is already in the reporting
# software update group because if it is it won't be added twice.
    Param(
        [Parameter(Mandatory = $true)]
        $SiteServerName,
        [Parameter(Mandatory = $true)]
        $SiteCode,
        [Parameter(Mandatory = $true)]
        $SUGName,
        [Parameter(Mandatory = $false)]
        $updInSUG,
        [Parameter(Mandatory = $false)]
        $updNotInSUG
        )
    #finding updates that are in Rpt SUG but not in any Non-Rpt SUG
    $updToRemove = @()
    foreach ($update in $updInSUG){
        if (!($updNotInSUG -match $update)){
            $updToRemove += $update
            Write-Log -iTabs 5 "$update flagged for removal" -bConsole $true       
        }
    }
    #finding updates that aren't in Rpt SUG but are in any Non-Rpt SUG
    $updToAdd = @()
    foreach ($update in $updNotInSUG){
        if (!($updInSUG -match $update)){
            $updToAdd += $update      
            Write-Log -iTabs 5 "$update flagged for addition" -bConsole $true        
        }
    }
    #removing extra updates from SUG Rpt
    if ($updToRemove.Count -gt 0){        
        Write-Log -iTabs 4 "Removing $($updToRemove.Count) from $SUGName" -bConsole $true         
        if ($action -like "*Run"){
            Remove-CMSoftwareUpdateFromGroup -SoftwareUpdateId $updToRemove -SoftwareUpdateGroupName $SUGName -Force #-warningaction silentlycontinue
        }
        <#$updcnt=1
        foreach ($upd in $updToRemove){
            try{
                if ($action -like "*Run"){
                    Remove-CMSoftwareUpdateFromGroup -SoftwareUpdateId $upd -SoftwareUpdateGroupName $SUGName -Force -warningaction silentlycontinue
                }
                Write-Log -iTabs 5 "($updcnt/$($updToRemove.Count)) - Removed $upd from $SUGName" -bConsole $true
                $updcnt++
            }
            catch{
                Write-Log -iTabs 5 "Error while running Remove-CMSoftwareUpdateFromGroup" -bConsole $true -sColor Red                 
            }                
        }#>
        Write-Log -iTabs 5 "$($updToRemove.Count) updates removed from $SUGName" -bConsole $true -sColor Green                              
    }
    else{
        Write-Log -iTabs 4 "No updates to remove from $SUGName" -bConsole $true
    }
    #adding updates
    if ($updToAdd.Count -gt 0){        
        Write-Log -iTabs 4 "Adding $($updToAdd.Count) to $SUGName" -bConsole $true         
        if ($action -like "*Run"){
            Add-CMSoftwareUpdateToGroup -SoftwareUpdateId $updToAdd -SoftwareUpdateGroupName $SUGName -Force #-warningaction silentlycontinue               
        }
        <#
        $updcnt=1

        foreach ($upd in $updToAdd){ 
            try{       
                if ($action -like "*Run"){     
                    Add-CMSoftwareUpdateToGroup -SoftwareUpdateId $upd -SoftwareUpdateGroupName $SUGName -Force -warningaction silentlycontinue               
                }
                Write-Log -iTabs 5 "($updcnt/$($updToAdd.Count)) - Added $upd to $SUGName" -bConsole $true
                $updcnt++
            }
            catch{
                Write-Log -iTabs 5 "Error while running Add-CMSoftwareUpdateToGroup" -bConsole $true -sColor Red 
            }                
        }#>
        Write-Log -iTabs 5 "$($updToAdd.Count) updates added to $SUGName" -bConsole $true -sColor Green                 
    }
    else{
        Write-Log -iTabs 4 "No updates to add to $SUGName" -bConsole $true
    }
}
function Set-SUGPair{

    Param(
        [Parameter(Mandatory = $true)]
        $SiteProviderServerName,
        [Parameter(Mandatory = $true)]
        $SiteCode,
        [Parameter(Mandatory = $true)]
        $CurrentUpdateGroup,
        [Parameter(Mandatory = $true)]
        $CurUpdList,        
        $PersistentUpdateGroup,        
        $PerUpdList,
        [Parameter(Mandatory = $false)]  
        $HandleAgedUpdates=$false,               
        $aAgedUpdates, 
        [Parameter(Mandatory = $false)]
        $PurgeExpired=$false,        
        $aExpUpdates,
        [Parameter(Mandatory = $false)]
        $PurgeSuperseded=$false,
        $aSupersededUpdates,
        [Parameter(Mandatory = $false)]
        $pkgSusName,
        [Parameter(Mandatory = $false)]
        $pkgSusList=$false
        )
    # If Current and persistent SUGs are equal, exit
    If ($CurrentUpdateGroup -eq $PersistentUpdateGroup){
        write-host ("The Current and Persistent update groups are the same group.  This is not allowed.  Exiting")
        exit
    }         
    #starting arrays
    $updatesToRemove =@()
    $updatesToMove   =@()

    ForEach ($UpdateID in $CurUpdList){               
        If (($PurgeExpired) -and (Test-SccmUpdateExpired -UpdateID $UpdateID -ExpUpdates $aExpUpdates)){            
            Write-Log -iTabs 4 "(CI_ID:$UpdateId) Expired." -bConsole $true -sColor DarkGray
            $updatesToRemove += $updateID            
        }
        elseIf (($PurgeSuperseded) -and (Test-SCCMUpdateSuperseded -UpdateID $UpdateID -SupUpdate $aSupersededUpdates)){            
            Write-Log -iTabs 4 "(CI_ID:$UpdateId) Superseded." -bConsole $true -sColor DarkYellow
            $updatesToRemove += $updateID
        }
        elseIf (($HandleAgedUpdates) -and (Test-SCCMUpdateAge -UpdateID $UpdateID -OldUpdates $aAgedUpdates)){
            Write-Log -iTabs 4 "(CI_ID:$UpdateId) Aged." -bConsole $true -sColor DarkGreen
            $updatesToMove += $updateID
        }
        else{
            #Write-Log -iTabs 4 "(CI_ID:$UpdateId) valid." -bConsole $true
        }
    }
    #If Superseded or Expired updates were flagged, script will remove them now
    If ($updatesToRemove.Count -gt 0){
        Write-Log -iTabs 4 "Removing $($updatesToRemove.Count) updates from $CurrentUpdateGroup due to being Expired or Superseded" -bConsole $true         
        try{
            $updcnt=1            
            foreach ($upd in $updatesToRemove){
                if ($action -like "*Run"){ 
                    Remove-CMSoftwareUpdateFromGroup -SoftwareUpdateId $upd -SoftwareUpdateGroupName $CurrentUpdateGroup -Force
                }
                Write-Log -iTabs 5 "($updcnt/$($updatesToRemove.Count)) - Update $upd removed from $CurrentUpdateGroup" -bConsole $true
                $updcnt++
            }
            Write-Log -iTabs 5 "All $($updatesToRemove.Count) flagged updates removed from $CurrentUpdateGroup" -bConsole $true -sColor Green
        }
        catch{
            Write-Log -iTabs 5 "Error while running Remove-CMSoftwareUpdateFromGroup" -bConsole $true -sColor Red 
        }
    }        
    #If aged updates were flagged, script will check if they need to be downloaded to sustainer, add them to sustainer SUG and finally remove from current SUG
    If (($updatesToMove.Count -gt 0) -and ($HandleAgedUpdates)){
        # due to the sensitivity of this process, a controll variable will be used named "$proceed" once it enters any major loop it will get value as $false.
        # once major loop is completed without errors the variable will receive value $true, allowing next major loop to start
        Write-Log -iTabs 4 "Adding $($updatesToMove.Count) updates to $PersistentUpdateGroup due to being Aged" -bConsole $true            
        # checking if there is a need to download updates
        Write-Log -iTabs 5 "Checking if updates to be moved, have to be downloaded." -bConsole $true
        $updatesToDownload =@()
        $downloadUpd=$false
        foreach ($update in $updatesToMove){            
            if (!($pkgSusList.CI_ID -match $update)){
                $updatesToDownload += $update
                $downloadUpd=$true
            }
        }
        Write-Log -iTabs 5 "Found $($updatesToDownload.Count) updates to be downloaded." -bConsole $true
        # downloading updates if needed
        $proceed=$true
        if (($downloadUpd) -and ($proceed)){            
            $proceed=$false
            Write-Log -iTabs 5 "Downloading $($updatesToDownload.Count) updates." -bConsole $true
            $updcnt=0
            foreach ($upd in $updatesToDownload){
                try{                    
                    if ($action -like "*Run"){
                        Save-CMSoftwareUpdate -SoftwareUpdateId $upd -DeploymentPackageName $pkgSusName -SoftwareUpdateLanguage "English" -DisableWildcardHandling -WarningAction SilentlyContinue                         
                    }
                    $updcnt++
                    Write-Log -iTabs 6 "($updcnt/$($updatesToDownload.Count)) - Update $upd downloaded to $pkgSusName pkg." -bConsole $true                    
                }
                catch{            
                    Write-Log -iTabs 6 "Error Downloading $upd into $pkgSusName." -bConsole $true -sColor red                                        
                    $global:iExitCode = 9015                     
                }
            }
            
            Write-Log -iTabs 6 "$($updatesToDownload.Count) updates Downloaded into $pkgSusName." -bConsole $true -sColor Green
            if ($updcnt -eq $updatesToDownload.count){
                $proceed=$true
            }
        }
        else{
            Write-Log -iTabs 5 "No need to download updates at this moment." -bConsole $true
        }
        # Adding updates to Sustainer
        if (!($proceed)){
            Write-Log -iTabs 4 "Failure detected while downloading KBs into Sustainer. KB Move will not proceed." -bConsole $true -sColor red
        }
        else{
            $proceed=$false
            try{            
                Write-Log -iTabs 5 "Adding $($updatesToMove.Count) to Sustainer SUG." -bConsole $true
                $updcnt=0
                foreach ($upd in $updatesToMove){
                    if ($action -like "*Run"){  
                        Add-CMSoftwareUpdateToGroup -SoftwareUpdateId $upd -SoftwareUpdateGroupName $PersistentUpdateGroup -Force -WarningAction SilentlyContinue
                    }
                    $updcnt++
                    Write-Log -iTabs 5 "($updcnt/$($updatesToMove.Count)) - $upd added to $PersistentUpdateGroup." -bConsole $true                    
                }                
                Write-Log -iTabs 5 "$($updatesToMove.Count) updates added to $PersistentUpdateGroup" -bConsole $true -sColor Green
            }
            catch{
                Write-Log -iTabs 5 "Error while running Add-CMSoftwareUpdateToGroup" -bConsole $true -sColor Red 
                Write-Log -iTabs 5 "Aborting script." -bConsole $true -sColor red
                $global:iExitCode = 9015
                return $global:iExitCode
            }
            if ($updcnt -eq $updatesToMove.Count){
                $proceed=$true
            }  
        }              
        # removing updates from Monthly SUG
        if ($proceed){
            Write-Log -iTabs 4 "Removing $($updatesToMove.Count) from $CurrentUpdateGroup due to being Aged" -bConsole $true            
            $updcnt=1
            foreach ($upd in $updatesToMove){
                try{
                    if ($action -like "*Run"){  
                        Remove-CMSoftwareUpdateFromGroup -SoftwareUpdateId $upd -SoftwareUpdateGroupName $CurrentUpdateGroup -Force
                    }
                    Write-Log -iTabs 5 "($updcnt/$($updatesToMove.Count)) - $upd added to $PersistentUpdateGroup." -bConsole $true
                    $updcnt++
                }
                catch{
                    Write-Log -iTabs 4 "Error while running Remove-CMSoftwareUpdateFromGroup" -bConsole $true -sColor Red                 
                }                
            }    
            Write-Log -iTabs 4 "$($updatesToMove.Count) updates removed from $CurrentUpdateGroup" -bConsole $true -sColor Green          
            #if updates removed and moved adds up to total updates, delete SUG
            if ($CurUpdList.count -eq ($updatesToMove.Count+$updatesToRemove.Count)){
                Write-Log -iTabs 4 "No updates left in $CurrentUpdateGroup. SUG will be deleted." -bConsole $true                    
                try{
                    if ($action -like "*Run"){
                        Remove-CMSoftwareUpdateGroup -Name $CurrentUpdateGroup -Force
                    }
                    Write-Log -iTabs 5 "SUG was deleted" -bConsole $true           
                }
                catch{
                    Write-Log -iTabs 5 "Error while deleting SUG." -bConsole $true -sColor Red                
                }
            }
        }
        else{
            Write-Log -iTabs 4 "Script will not remove updates from $CurrentUpdateGroup since it failed to add updates in Sustainer" -bConsole $true            
        }
    }
}
function Set-DeploymentPackages {
    Param(
        [Parameter(Mandatory = $false)]
        $SiteProviderServerName,
        [Parameter(Mandatory = $false)]
        $SiteCode,
        [Parameter(Mandatory = $false)]
        $monUpdList,
        [Parameter(Mandatory = $false)]
        $susUpdList,
        [Parameter(Mandatory = $false)]
        $pkgMonthlyList,
        [Parameter(Mandatory = $false)]
        $pkgSustainerList,
        [Parameter(Mandatory = $false)]
        $pkgMonthly,
        [Parameter(Mandatory = $false)]
        $pkgSustainer
        )   
    # Checkig if all Upd from SUGs are present in at least 1 pkg
    $updatesToDownloadMonth =@()
    $updatesToDownloadSus =@()
    Write-Log -iTabs 3 "Evaluating if downloaded updates are deployed in SUGs" -bConsole $true
    foreach ($update in $monUpdList){
        if (!($pkgMonthlyList.CI_ID -match $update)){
            $updatesToDownloadMonth += $update
        }
    }
    foreach ($update in $susUpdList){
        if (!($pkgSustainerList.CI_ID -match $update)){
            $updatesToDownloadSus += $update
        }
    }
    # Checking if all updates in Sustainer package is present in SUGs
    $updatesToDeleteSus = @()
    foreach ($update in $pkgSustainerList){
        if (!($susUpdList -match $update.CI_ID)){
            $updatesToDeleteSus += $update        
        }
    }
    $updatesToDeleteMonth = @()
    foreach ($update in $pkgMonthlyList){
        if (!($monUpdList -match $update.CI_ID)){
            $updatesToDeleteMonth += $update        
        }
    }
    # Deleting Updates from Sustainer package, if needed
    if ($updatesToDeleteSus.count -gt 0){
        Write-Log -iTabs 4 "Found $($updatesToDeleteSus.count) extra updates to be deleted from Aged Pkg" -bConsole $true
        #creating obj with Package WMI
        $susPackageWMI = Get-WmiObject -ComputerName $SiteProviderServerName -Namespace root\sms\site_$($SiteCode) -Class SMS_SoftwareUpdatesPackage -Filter "Name ='$pkgSustainer'"
        $ContentIDArray = @()
        # Converting CI_ID into ContentID
        foreach($upd in $updatesToDeleteSus){            
            $ContentIDtoContent = Get-WMIObject -NameSpace root\sms\site_$($SiteCode) -Class SMS_CItoContent -Filter "CI_ID='$($upd.CI_ID)'"
            $ContentIDArray += $ContentIDtoContent.ContentID
        }
        Write-Log -iTabs 5 "Converted updates into $($contentIDArray.count) ContentIDs" -bConsole $true
        #calling Remove Content Method known WMI bug might cause a temporary failure. Adding loop for resiliency while removing content.        
        $pause=0
        $pkgClean = $false
        while($pkgClean -eq $false){
            try{                   
                Start-Sleep $pause
                if ($action -like "*Run"){
                    $susPackageWMI.RemoveContent($ContentIDArray,$true) | Out-Null
                }
                $pkgClean=$true
                Write-Log -itabs 5 "Package clean-up finished" -bConsole $true
            }
            catch{                             
                Write-Log -itabs 5 "Package clean-up failed, but is a known issue. Will try again" -bConsole $true -sColor red               
                $pause +=5                
                if ($pause -eq 25 ){
                    $pkgClean=$true
                    Write-Log -itabs 5 "Unable to clena package. Try again later" -bConsole $true -sColor red               
                }
            }
        }       
    }    
    # Deleting Updates from Monthly package, if needed
    if ($updatesToDeleteMonth.count -gt 0){
        Write-Log -iTabs 4 "Found $($updatesToDeleteMonth.count) extra updates to be deleted from Monthly Pkg" -bConsole $true
        #creating obj with Package WMI
        $monPackageWMI = Get-WmiObject -ComputerName $SiteProviderServerName -Namespace root\sms\site_$($SiteCode) -Class SMS_SoftwareUpdatesPackage -Filter "Name ='$pkgMonthly'"
        $ContentIDArray = @()
        # Converting CI_ID into ContentID
        foreach($upd in $updatesToDeleteMonth){            
            $ContentIDtoContent = Get-WMIObject -NameSpace root\sms\site_$($SiteCode) -Class SMS_CItoContent -Filter "CI_ID='$($upd.CI_ID)'"
            $ContentIDArray += $ContentIDtoContent.ContentID
        }
        Write-Log -iTabs 5 "Converted updates into $($contentIDArray.count) ContentIDs" -bConsole $true
        #calling Remove Content Method known WMI bug might cause a temporary failure. Adding loop for resiliency while removing content.
        $pause=0
        $pkgClean = $false
        while($pkgClean -eq $false){
            try{                   
                Start-Sleep $pause
                if ($action -like "*Run"){
                    $MonPackageWMI.RemoveContent($ContentIDArray,$true) | Out-Null
                }
                $pkgClean=$true
                Write-Log -itabs 5 "Package clean-up finished" -bConsole $true
            }
            catch{                             
                Write-Log -itabs 5 "Package clean-up failed, but is a known issue. Will try again" -bConsole $true -sColor red               
                $pause +=5                
                if ($pause -eq 25 ){
                    $pkgClean=$true
                    Write-Log -itabs 5 "Unable to clena package. Try again later" -bConsole $true -sColor red               
                }
            }
        }
    }    
    # Downloading updates to Sustainer package, if needed
    if ($updatesToDownloadSus.count -gt 0){
        Write-Log -iTabs 4 "Found $($updatesToDownloadSus.count) required to be downloaded into Sustainer Pkg" -bConsole $true
        $updcnt=1
        Foreach ($upd in $updatesToDownloadSus){
            try{
                if ($action -like "*Run"){                
                    Save-CMSoftwareUpdate -SoftwareUpdateId $upd -DeploymentPackageName $pkgSustainer -SoftwareUpdateLanguage "English" -DisableWildcardHandling -WarningAction SilentlyContinue                         
                }
                Write-Log -iTabs 5 "($updcnt/$($updatesToDownloadSus.count)) - Update $upd downloaded to Sustainer Pkg." -bConsole $true
                $updcnt++
            }
            catch{
                Write-Log -iTabs 5 "$updcnt - Error Downloading $upd into Sustainer Pkg." -bConsole $true -sColor red                                                        
                $updcnt++
            }
        }
    }
    # Downloading Updates to Monthly Package, if needed
    if ($updatesToDownloadMonth.count -gt 0){
        Write-Log -iTabs 4 "Found $($updatesToDownloadMonth.count) required to be downloaded into Monthly Pkg" -bConsole $true
        $updcnt=1
        Foreach ($upd in $updatesToDownloadMonth){
            try{
                if ($action -like "*Run"){                
                    Save-CMSoftwareUpdate -SoftwareUpdateId $upd -DeploymentPackageName $pkgMonthly -SoftwareUpdateLanguage "English" -DisableWildcardHandling -WarningAction SilentlyContinue                         
                }
                Write-Log -iTabs 5 "($updcnt/$($updatesToDownloadMonth.count)) - Update $upd downloaded to Monthly Pkg." -bConsole $true
                $updcnt++
            }
            catch{
                Write-Log -iTabs 5 "$updcnt - Error Downloading $upd into Monthly Pkg." -bConsole $true -sColor red                                        
                $global:iExitCode = 9015
                $updcnt++
            }
        }
    }
    Write-Log -iTabs 3 "Deployment Packages review is now complete"

}
function Get-NumUpdInGroups{
# This script will examine the count of updates in each deployed update group and provide a warning
# when the number of updates in a given group exceeds 900.
    Param(
        [Parameter(Mandatory = $true)]
        $SiteServerName,
        [Parameter(Mandatory = $true)]
        $SiteCode,
        [Parameter(Mandatory = $true)]
        $sugs
        )    
    # Loop through each software update group and check the total number of updates in each.    
    ForEach ($sug in $sugs | Sort-Object $sugs.LocalizedDisplayName){        
        # Only test update groups that are deployed.  Reporting software update groups may be used
        # in some environments and as long as these groups aren't deployed they can contain greater
        # than 1000 updates.  Accordingly, warning for those groups doesn't apply.
        if (($sug.Updates.Count -gt 900) -and ($sug.IsDeployed -eq 'True')){
            $textColor="Red"
        }
        else{
            $textColor="white"
        }           
        write-log -itabs 4 "# of Updates found in $($sug.LocalizedDisplayName): $($sug.Updates.Count)." -bConsole $true -sColor $textColor
        if ($textcolor -eq "Red"){            
            write-log -itabs 5 "SUGs deployed should contain less than 900 updates. Consider splitting this SUG into more." -bConsole $true -sColor $textColor
        }
        if (($sug.Updates.Count -eq 0) -and ($sug.DateCreated -lt $timeSustainerAge)){
            write-log -itabs 5 "Empty SUGs older than Aged threshold.Deleting it..." -bConsole $true
            try{
                Remove-CMSoftwareUpdateGroup -Name $sug.LocalizedDisplayName    
                write-log -itabs 5 "Deleted!" -bConsole $true -sColor Green
            }
            catch{
                write-log -itabs 5 "Error while deleting $($sug.LocalizedDisplayName) SUG!" -bConsole $true -sColor Yellow -bEventLog $true -iEventID 7001 -sSource $sEventSource
            }
        }
    }     
}
function Delete-OldDeployments{
    Param(
        $SiteServerName,
        $SiteCode,               
        $sugName,
        $sugID,
        $CollectiontemplateName)   
    #list all deployments
        Write-Log -iTabs 4 "Getting all deployments from Software Update Group" -bConsole $true
        $deployments = Get-CMUpdateGroupDeployment | Where-Object {$_.AssignedUpdateGroup -eq "$sugID"}
        $collectionIDs = Get-CMCollection -Name $CollectiontemplateName* | Select CollectionID
        foreach ($deployment in $deployments){
            if ($deployment.TargetCollectionID -notin $collectionIDs.CollectionID){
                Write-Log -iTabs 5 "$($deployment.AssignmentName) was found as old deployment" -bConsole $true
                try{
                    if ($action -like "*Run"){
                        Remove-CMUpdateGroupDeployment -DeploymentId $deployment.AssignmentUniqueID -Force
                    }
                    Write-Log -iTabs 6 "Deployment removed" -bConsole $true
                }
                catch{
                    Write-Log -iTabs 6 "Error while Deployment removed" -bConsole $true -sColor red
                }
            }
        }
        Write-Log -iTabs 4 "Deployment clean-up is complete" -bConsole $true
} 
#endregion
# --------------------------------------------------------------------------------------------
#region VARIABLES
# Standard Variables
    # *****  Change Logging Path and File Name Here  *****    
    $sOutFileName	= "Renew-ReportSUGs.log" # Log File Name    
    $sEventSource   = "ToolBox" # Event Source Name
    # ****************************************************
    $sScriptName 	= $MyInvocation.MyCommand
    $sScriptPath 	= Split-Path -Parent $MyInvocation.MyCommand.Path
    $sLogRoot		= Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\SMS\Tracing\
    $sLogRoot       = $sLogRoot[0].GetValue('Tracefilename')
    $sLogRoot       = $sLogRoot.Substring(0,$SLogRoot.LastIndexOf('\')+1)    
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
    $timeReport = 30 # Defines how long before an update is considered within Report Groups
    
    switch ($Scope){
        #IF CAS
        "CAS"{
            $SMSProvider = "sccm01.zlab.varandas.com"            
            $SCCMSite = "CAS"            
            $TemplateName = "WKS-SecurityUpdates-"                
            $timeReport = 45     
            $severity = 8
            $updateProducts = {
                'fa5ef799-b817-439e-abf7-c76ba0cacb75', #ASP.NET Web Frameworks
                '83aed513-c42d-4f94-b4dc-f2670973902d', #CAPICOM
                '0bbd2260-7478-4553-a791-21ab88e437d2', #Device Health
                '5e870422-bd8f-4fd2-96d3-9c5d9aafda22', #Microsoft Lync 2010
                'e9ece729-676d-4b57-b4d1-7e0ab0589707', #Microsoft SQL Server 2008 R2 - PowerPivot for Microsoft Excel 2010
                '56750722-19b4-4449-a547-5b68f19eee38', #Microsoft SQL Server 2012
                'caab596c-64f2-4aa9-bbe3-784c6e2ccf9c', #Microsoft SQL Server 2014
                '93f0b0bc-9c20-4ca5-b630-06eb4706a447', #Microsoft SQL Server 2016
                'ca6616aa-6310-4c2d-a6bf-cae700b85e86', #Microsoft SQL Server 2017
                'dee854fd-e9d2-43fd-bbc3-f7568e3ce324', #Microsoft SQL Server Management Studio v17
                '6b9e8b26-8f50-44b9-94c6-7846084383ec', #MS Security Essentials
                '6248b8b1-ffeb-dbd9-887a-2acf53b09dfe', #Office 2002/XP
                '1403f223-a63f-f572-82ba-c92391218055', #Office 2003
                '041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591', #Office 2007
                '84f5f325-30d7-41c4-81d1-87a0e6535b66', #Office 2010
                '704a0a4a-518f-4d69-9e03-10ba44198bd5', #Office 2013
                '25aed893-7c2d-4a31-ae22-28ff8ac150ed', #Office 2016
                '6c5f2e66-7dbc-4c59-90a7-849c4c649d7a', #Office 2019
                '30eb551c-6288-4716-9a78-f300ec36d72b', #Office 365 Client
                '8bc19572-a4b6-4910-b70d-716fecffc1eb', #Office Communicator 2007 R2
                '7cf56bdd-5b4e-4c04-a6a6-706a2199eff7', #Report Viewer 2005
                '79adaa30-d83b-4d9c-8afd-e099cf34855f', #Report Viewer 2008
                'f7f096c9-9293-422d-9be8-9f6e90c2e096', #Report Viewer 2010
                '6cf036b9-b546-4694-885a-938b93216b66', #Security Essentials
                '9f3dd20a-1004-470e-ba65-3dc62d982958', #Silverlight
                '7145181b-9556-4b11-b659-0162fa9df11f', #SQL Server 2000
                '60916385-7546-4e9b-836e-79d65e517bab', #SQL Server 2005
                'c5f0b23c-e990-4b71-9808-718d353f533a', #SQL Server 2008
                'bb7bc3a7-857b-49d4-8879-b639cf5e8c3c', #SQL Server 2008 R2
                '7fe4630a-0330-4b01-a5e6-a77c7ad34eb0', #SQL Server 2012 Product Updates for Setup
                '892c0584-8b03-428f-9a74-224fcd6887c0', #SQL Server 2014-2016 Product Updates for Setup
                'c96c35fc-a21f-481b-917c-10c4f64792cb', #SQL Server Feature Pack
                'bf05abfb-6388-4908-824e-01565b05e43a', #System Center 2012 - Operations Manager
                'ab8df9b9-8bff-4999-aee5-6e4054ead976', #System Center 2012 - Orchestrator
                '2a9170d5-3434-4820-885c-61a4f3fc6f84', #System Center 2012 R2 - Operations Manager
                '6ddf2e90-4b40-471c-a664-6cd6b7e0d0a7', #System Center 2012 R2 - Orchestrator
                'cc4ab3ac-9d9a-4f53-97d3-e0d6de39d119', #System Center 2016 - Operations Manager
                'e505a854-6941-484f-a107-ebf0d1b64820', #System Center 2016 - Orchestrator
                'cd8d80fe-5b55-48f1-b37a-96535dca6ae7', #TMG Firewall Client
                'a0dd7e72-90ec-41e3-b370-c86a245cd44f', #Visual Studio 2005
                'e3fde9f8-14d6-4b5c-911c-fba9e0fc9887', #Visual Studio 2008
                'c9834186-a976-472b-8384-6bb8f2aa43d9', #Visual Studio 2010
                'cbfd1e71-9d9e-457e-a8c5-500c47cfe9f3', #Visual Studio 2010 Tools for Office Runtime
                'e1c753f2-9f79-4577-b75b-913f4230feee', #Visual Studio 2010 Tools for Office Runtime
                'abddd523-04f4-4f8e-b76f-a6c84286cc67', #Visual Studio 2012
                'cf4aa0fc-119d-4408-bcba-181abb69ed33', #Visual Studio 2013
                '1731f839-8830-4b9c-986e-82ee30b24120', #Visual Studio 2015
                'a3c2375d-0c8a-42f9-bce0-28333e198407', #Windows 10
                'bfe5b177-a086-47a0-b102-097e4fa1f807', #Windows 7
                '6407468e-edc7-4ecd-8c32-521f64cee65e', #Windows 8.1
                'b1b8f641-1ff2-4ae6-b247-4fe7503787be', #Windows Admin Center
                '8c3fcc84-7410-4a95-8b89-a166a0190486', #Windows Defender
                '50c04525-9b15-4f7c-bed4-87455bcd7ded', #Windows Dictionary Updates
                'ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf', #Windows Server 2008
                'fdfe8200-9d98-44ba-a12a-772282bf60ef', #Windows Server 2008 R2
                'd31bd4c3-d872-41c9-a2e7-231f372588cb', #Windows Server 2012 R2
                '569e8e8f-c6cd-42c8-92a3-efbb20a0f6f5', #Windows Server 2016                
                'f702a48c-919b-45d6-9aef-ca4248d50397', #Windows Server 2019
                '4e487029-f550-4c22-8b31-9173f3f95786', #Windows Server Manager – Windows Server Update Services (WSUS) Dynamic Installer               
                '26997d30-08ce-4f25-b2de-699c36a8033a' #Windows Vista
            }
            $updateClassification = '0fa1201d-4330-4fa8-8ae9-b877473b6441' #Security Updates   
        }          
        #IF VAR
        "VAR"{
            $SMSProvider = "sccm01.vlab.varandas.com"
            $SCCMSite = "VAR"
            $TemplateName = "VAR-"                        
            $timeReport = 15
            $severity = 8
            $updateProducts = {
                'fa5ef799-b817-439e-abf7-c76ba0cacb75', #ASP.NET Web Frameworks
                '83aed513-c42d-4f94-b4dc-f2670973902d', #CAPICOM
                '0bbd2260-7478-4553-a791-21ab88e437d2', #Device Health
                '5e870422-bd8f-4fd2-96d3-9c5d9aafda22', #Microsoft Lync 2010
                'e9ece729-676d-4b57-b4d1-7e0ab0589707', #Microsoft SQL Server 2008 R2 - PowerPivot for Microsoft Excel 2010
                '56750722-19b4-4449-a547-5b68f19eee38', #Microsoft SQL Server 2012
                'caab596c-64f2-4aa9-bbe3-784c6e2ccf9c', #Microsoft SQL Server 2014
                '93f0b0bc-9c20-4ca5-b630-06eb4706a447', #Microsoft SQL Server 2016
                'ca6616aa-6310-4c2d-a6bf-cae700b85e86', #Microsoft SQL Server 2017
                'dee854fd-e9d2-43fd-bbc3-f7568e3ce324', #Microsoft SQL Server Management Studio v17
                '6b9e8b26-8f50-44b9-94c6-7846084383ec', #MS Security Essentials
                '6248b8b1-ffeb-dbd9-887a-2acf53b09dfe', #Office 2002/XP
                '1403f223-a63f-f572-82ba-c92391218055', #Office 2003
                '041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591', #Office 2007
                '84f5f325-30d7-41c4-81d1-87a0e6535b66', #Office 2010
                '704a0a4a-518f-4d69-9e03-10ba44198bd5', #Office 2013
                '25aed893-7c2d-4a31-ae22-28ff8ac150ed', #Office 2016
                '6c5f2e66-7dbc-4c59-90a7-849c4c649d7a', #Office 2019
                '30eb551c-6288-4716-9a78-f300ec36d72b', #Office 365 Client
                '8bc19572-a4b6-4910-b70d-716fecffc1eb', #Office Communicator 2007 R2
                '7cf56bdd-5b4e-4c04-a6a6-706a2199eff7', #Report Viewer 2005
                '79adaa30-d83b-4d9c-8afd-e099cf34855f', #Report Viewer 2008
                'f7f096c9-9293-422d-9be8-9f6e90c2e096', #Report Viewer 2010
                '6cf036b9-b546-4694-885a-938b93216b66', #Security Essentials
                '9f3dd20a-1004-470e-ba65-3dc62d982958', #Silverlight
                '7145181b-9556-4b11-b659-0162fa9df11f', #SQL Server 2000
                '60916385-7546-4e9b-836e-79d65e517bab', #SQL Server 2005
                'c5f0b23c-e990-4b71-9808-718d353f533a', #SQL Server 2008
                'bb7bc3a7-857b-49d4-8879-b639cf5e8c3c', #SQL Server 2008 R2
                '7fe4630a-0330-4b01-a5e6-a77c7ad34eb0', #SQL Server 2012 Product Updates for Setup
                '892c0584-8b03-428f-9a74-224fcd6887c0', #SQL Server 2014-2016 Product Updates for Setup
                'c96c35fc-a21f-481b-917c-10c4f64792cb', #SQL Server Feature Pack
                'bf05abfb-6388-4908-824e-01565b05e43a', #System Center 2012 - Operations Manager
                'ab8df9b9-8bff-4999-aee5-6e4054ead976', #System Center 2012 - Orchestrator
                '2a9170d5-3434-4820-885c-61a4f3fc6f84', #System Center 2012 R2 - Operations Manager
                '6ddf2e90-4b40-471c-a664-6cd6b7e0d0a7', #System Center 2012 R2 - Orchestrator
                'cc4ab3ac-9d9a-4f53-97d3-e0d6de39d119', #System Center 2016 - Operations Manager
                'e505a854-6941-484f-a107-ebf0d1b64820', #System Center 2016 - Orchestrator
                'cd8d80fe-5b55-48f1-b37a-96535dca6ae7', #TMG Firewall Client
                'a0dd7e72-90ec-41e3-b370-c86a245cd44f', #Visual Studio 2005
                'e3fde9f8-14d6-4b5c-911c-fba9e0fc9887', #Visual Studio 2008
                'c9834186-a976-472b-8384-6bb8f2aa43d9', #Visual Studio 2010
                'cbfd1e71-9d9e-457e-a8c5-500c47cfe9f3', #Visual Studio 2010 Tools for Office Runtime
                'e1c753f2-9f79-4577-b75b-913f4230feee', #Visual Studio 2010 Tools for Office Runtime
                'abddd523-04f4-4f8e-b76f-a6c84286cc67', #Visual Studio 2012
                'cf4aa0fc-119d-4408-bcba-181abb69ed33', #Visual Studio 2013
                '1731f839-8830-4b9c-986e-82ee30b24120', #Visual Studio 2015
                'a3c2375d-0c8a-42f9-bce0-28333e198407', #Windows 10
                'bfe5b177-a086-47a0-b102-097e4fa1f807', #Windows 7
                '6407468e-edc7-4ecd-8c32-521f64cee65e', #Windows 8.1
                'b1b8f641-1ff2-4ae6-b247-4fe7503787be', #Windows Admin Center
                '8c3fcc84-7410-4a95-8b89-a166a0190486', #Windows Defender
                '50c04525-9b15-4f7c-bed4-87455bcd7ded', #Windows Dictionary Updates
                'ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf', #Windows Server 2008
                'fdfe8200-9d98-44ba-a12a-772282bf60ef', #Windows Server 2008 R2
                'd31bd4c3-d872-41c9-a2e7-231f372588cb', #Windows Server 2012 R2
                '569e8e8f-c6cd-42c8-92a3-efbb20a0f6f5', #Windows Server 2016                
                'f702a48c-919b-45d6-9aef-ca4248d50397', #Windows Server 2019
                '4e487029-f550-4c22-8b31-9173f3f95786', #Windows Server Manager – Windows Server Update Services (WSUS) Dynamic Installer               
                '26997d30-08ce-4f25-b2de-699c36a8033a' #Windows Vista
            }
            $updateClassification = '0fa1201d-4330-4fa8-8ae9-b877473b6441' #Security Updates                          
        }       
        #IF PVA
        "PVA"{
            $SMSProvider = "sccm01.plab.varandas.com"
            $SCCMSite = "PVA"            
            $TemplateName = "VAR-"                           
            $timeReport = 20     
            $severity = 8
            $updateProducts = {
                'fa5ef799-b817-439e-abf7-c76ba0cacb75', #ASP.NET Web Frameworks
                '83aed513-c42d-4f94-b4dc-f2670973902d', #CAPICOM
                '0bbd2260-7478-4553-a791-21ab88e437d2', #Device Health
                '5e870422-bd8f-4fd2-96d3-9c5d9aafda22', #Microsoft Lync 2010
                'e9ece729-676d-4b57-b4d1-7e0ab0589707', #Microsoft SQL Server 2008 R2 - PowerPivot for Microsoft Excel 2010
                '56750722-19b4-4449-a547-5b68f19eee38', #Microsoft SQL Server 2012
                'caab596c-64f2-4aa9-bbe3-784c6e2ccf9c', #Microsoft SQL Server 2014
                '93f0b0bc-9c20-4ca5-b630-06eb4706a447', #Microsoft SQL Server 2016
                'ca6616aa-6310-4c2d-a6bf-cae700b85e86', #Microsoft SQL Server 2017
                'dee854fd-e9d2-43fd-bbc3-f7568e3ce324', #Microsoft SQL Server Management Studio v17
                '6b9e8b26-8f50-44b9-94c6-7846084383ec', #MS Security Essentials
                '6248b8b1-ffeb-dbd9-887a-2acf53b09dfe', #Office 2002/XP
                '1403f223-a63f-f572-82ba-c92391218055', #Office 2003
                '041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591', #Office 2007
                '84f5f325-30d7-41c4-81d1-87a0e6535b66', #Office 2010
                '704a0a4a-518f-4d69-9e03-10ba44198bd5', #Office 2013
                '25aed893-7c2d-4a31-ae22-28ff8ac150ed', #Office 2016
                '6c5f2e66-7dbc-4c59-90a7-849c4c649d7a', #Office 2019
                '30eb551c-6288-4716-9a78-f300ec36d72b', #Office 365 Client
                '8bc19572-a4b6-4910-b70d-716fecffc1eb', #Office Communicator 2007 R2
                '7cf56bdd-5b4e-4c04-a6a6-706a2199eff7', #Report Viewer 2005
                '79adaa30-d83b-4d9c-8afd-e099cf34855f', #Report Viewer 2008
                'f7f096c9-9293-422d-9be8-9f6e90c2e096', #Report Viewer 2010
                '6cf036b9-b546-4694-885a-938b93216b66', #Security Essentials
                '9f3dd20a-1004-470e-ba65-3dc62d982958', #Silverlight
                '7145181b-9556-4b11-b659-0162fa9df11f', #SQL Server 2000
                '60916385-7546-4e9b-836e-79d65e517bab', #SQL Server 2005
                'c5f0b23c-e990-4b71-9808-718d353f533a', #SQL Server 2008
                'bb7bc3a7-857b-49d4-8879-b639cf5e8c3c', #SQL Server 2008 R2
                '7fe4630a-0330-4b01-a5e6-a77c7ad34eb0', #SQL Server 2012 Product Updates for Setup
                '892c0584-8b03-428f-9a74-224fcd6887c0', #SQL Server 2014-2016 Product Updates for Setup
                'c96c35fc-a21f-481b-917c-10c4f64792cb', #SQL Server Feature Pack
                'bf05abfb-6388-4908-824e-01565b05e43a', #System Center 2012 - Operations Manager
                'ab8df9b9-8bff-4999-aee5-6e4054ead976', #System Center 2012 - Orchestrator
                '2a9170d5-3434-4820-885c-61a4f3fc6f84', #System Center 2012 R2 - Operations Manager
                '6ddf2e90-4b40-471c-a664-6cd6b7e0d0a7', #System Center 2012 R2 - Orchestrator
                'cc4ab3ac-9d9a-4f53-97d3-e0d6de39d119', #System Center 2016 - Operations Manager
                'e505a854-6941-484f-a107-ebf0d1b64820', #System Center 2016 - Orchestrator
                'cd8d80fe-5b55-48f1-b37a-96535dca6ae7', #TMG Firewall Client
                'a0dd7e72-90ec-41e3-b370-c86a245cd44f', #Visual Studio 2005
                'e3fde9f8-14d6-4b5c-911c-fba9e0fc9887', #Visual Studio 2008
                'c9834186-a976-472b-8384-6bb8f2aa43d9', #Visual Studio 2010
                'cbfd1e71-9d9e-457e-a8c5-500c47cfe9f3', #Visual Studio 2010 Tools for Office Runtime
                'e1c753f2-9f79-4577-b75b-913f4230feee', #Visual Studio 2010 Tools for Office Runtime
                'abddd523-04f4-4f8e-b76f-a6c84286cc67', #Visual Studio 2012
                'cf4aa0fc-119d-4408-bcba-181abb69ed33', #Visual Studio 2013
                '1731f839-8830-4b9c-986e-82ee30b24120', #Visual Studio 2015
                'a3c2375d-0c8a-42f9-bce0-28333e198407', #Windows 10
                'bfe5b177-a086-47a0-b102-097e4fa1f807', #Windows 7
                '6407468e-edc7-4ecd-8c32-521f64cee65e', #Windows 8.1
                'b1b8f641-1ff2-4ae6-b247-4fe7503787be', #Windows Admin Center
                '8c3fcc84-7410-4a95-8b89-a166a0190486', #Windows Defender
                '50c04525-9b15-4f7c-bed4-87455bcd7ded', #Windows Dictionary Updates
                'ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf', #Windows Server 2008
                'fdfe8200-9d98-44ba-a12a-772282bf60ef', #Windows Server 2008 R2
                'd31bd4c3-d872-41c9-a2e7-231f372588cb', #Windows Server 2012 R2
                '569e8e8f-c6cd-42c8-92a3-efbb20a0f6f5', #Windows Server 2016                
                'f702a48c-919b-45d6-9aef-ca4248d50397', #Windows Server 2019
                '4e487029-f550-4c22-8b31-9173f3f95786', #Windows Server Manager – Windows Server Update Services (WSUS) Dynamic Installer               
                '26997d30-08ce-4f25-b2de-699c36a8033a' #Windows Vista
            }
            $updateClassification = '0fa1201d-4330-4fa8-8ae9-b877473b6441' #Security Updates                           
        }        
        default{
            $SMSProvider,$SCCMSite,$TemplateName = $null
            $severity=8
            $updateProducts = {
                #'fdcfda10-5b1f-4e57-8298-c744257e30db', #Active Directory Rights Management Services Client 2.0
                #'5d6a452a-55ba-4e11-adac-85e180bda3d6', #Antigen for Exchange/SMTP
                #'fa5ef799-b817-439e-abf7-c76ba0cacb75', #ASP.NET Web Frameworks
                #'fb08c71c-dbe9-40ab-8302-fb0231b1c814', #Azure File Sync agent updates for Windows Server 2012 R2
                #'7ff1d901-fd38-441b-aaba-36d7b0ebf264', #Azure File Sync agent updates for Windows Server 2016
                #'84a044f8-631c-4eb5-90be-9f1d127d6cc2', #Azure File Sync agent updates for Windows Server 2019
                #'b86cf33d-92ac-43d2-886b-be8a12f81ee1', #Bing Bar
                #'5349cd30-1ffd-731f-7f94-52c6774f2534', #Bios
                #'34aae785-2ae3-446d-b305-aec3770edcef', #BizTalk Server 2002
                #'86b9f801-b8ec-4d16-b334-08fba8567c17', #BizTalk Server 2006R2
                #'b61793e6-3539-4dc8-8160-df71054ea826', #BizTalk Server 2009
                #'61487ade-9a4e-47c9-baa5-f1595bcdc5c5', #BizTalk Server 2013
                #'83aed513-c42d-4f94-b4dc-f2670973902d', #CAPICOM
                #'236c566b-aaa6-482c-89a6-1e6c5cac6ed8', #Category for System Center Online Client
                #'ac615cb5-1c12-44be-a262-fab9cd8bf523', #Compute Cluster Pack
                #'eb658c03-7d9f-4bfa-8ef3-c113b7466e73', #Data Protection Manager 2006
                #'0bbd2260-7478-4553-a791-21ab88e437d2', #Device Health
                #'f76b7f51-b762-4fd0-a35c-e04f582acf42', #Dictionary Updates for Microsoft IMEs
                #'5a0031d6-edef-f08b-6a12-ff17ac03525e', #Drivers and Applications
                #'83a83e29-7d55-44a0-afed-aea164bc35e6', #Exchange 2000 Server
                #'3cf32f7c-d8ee-43f8-a0da-8b88a6f8af1a', #Exchange Server 2003
                #'26bb6be1-37d1-4ca6-baee-ec00b2f7d0f1', #Exchange Server 2007
                #'ab62c5bd-5539-49f6-8aea-5a114dd42314', #Exchange Server 2007 and Above Anti-spam
                #'9b135dd5-fc75-4609-a6ae-fb5d22333ef0', #Exchange Server 2010
                #'d3d7c7a6-3e2f-4029-85bf-b59796b82ce7', #Exchange Server 2013
                #'49c3ddde-4df2-4534-980c-83f4e27b23b5', #Exchange Server 2016
                #'f7fcd7d7-a163-4b27-970a-48bc02023df1', #Exchange Server 2019
                #'fa9ff215-cfe0-4d57-8640-c65f24e6d8e0', #Expression Design 1
                #'f3b1d39b-6871-4b51-8b8c-6eb556c8eee1', #Expression Design 2
                #'18a2cff8-9fd2-487e-ac3b-f490e6a01b2d', #Expression Design 3
                #'9119fae9-3fdd-4c06-bde7-2cbbe2cf3964', #Expression Design 4
                #'5108d510-e169-420c-9a4d-618bdb33c480', #Expression Media 2
                #'d8584b2b-3ac5-4201-91cb-caf6d240dc0b', #Expression Media V1
                #'a33f42ac-b33f-4fd2-80a8-78b3bfa6a142', #Expression Web 3
                #'3b1e1746-d99b-42d4-91fd-71d794f97a4d', #Expression Web 4
                #'d72155f3-8aa8-4bf7-9972-0a696875b74e', #Firewall Client for ISA Server
                #'0a487050-8b0f-4f81-b401-be4ceacd61cd', #Forefront Client Security
                #'a38c835c-2950-4e87-86cc-6911a52c34a3', #Forefront Endpoint Protection 2010
                #'d7d32245-1064-4edf-bd09-0218cfb6a2da', #Forefront Identity Manager 2010
                #'86134b1c-cf56-4884-87bf-5c9fe9eb526f', #Forefront Identity Manager 2010 R2
                #'a6432e15-a446-44af-8f96-0475c472aef6', #Forefront Protection Category
                #'f54d8a80-c7e1-476c-9995-3d6aee4bfb58', #Forefront Server Security Category
                #'84a54ea9-e574-457a-a750-17164c1d1679', #Forefront Threat Management Gateway, Definition Updates for HTTP Malware Inspection
                #'59f07fb7-a6a1-4444-a9a9-fb4b80138c6d', #Forefront TMG
                #'06bdf56c-1360-4bb9-8997-6d67b318467c', #Forefront TMG MBE
                #'2e068336-2ead-427a-b80d-5b0fffded7e7', #HealthVault Connection Center
                #'d84d138e-8423-4102-b317-91b1339aa9c9', #HealthVault Connection Center Upgrades
                #'0c6af366-17fb-4125-a441-be87992b953a', #Host Integration Server 2000
                #'784c9f6d-959a-433f-b7a3-b2ace1489a18', #Host Integration Server 2004
                #'eac7e88b-d8d4-4158-a828-c8fc1325a816', #Host Integration Server 2006
                #'42b678ae-2b57-4251-ae57-efbd35e7ae96', #Host Integration Server 2009
                #'3f3b071e-c4a6-4bcc-b6c1-27122b235949', #Host Integration Server 2010
                #'5964c9f1-8e72-4891-a03a-2aed1c4115d2', #HPC Pack 2008
                #'b627a8ff-19cd-45f5-a938-32879dd90123', #Internet Security and Acceleration Server 2004
                #'2cdbfa44-e2cb-4455-b334-fce74ded8eda', #Internet Security and Acceleration Server 2006
                #'5cc25303-143f-40f3-a2ff-803a1db69955', #Locally published packages
                #'5669bd12-c6ab-4831-8643-0d5f6638228f', #Max
                #'6ac905a5-286b-43eb-97e2-e23b3848c87d', #Microsoft Advanced Threat Analytics
                #'00b2d754-4512-4278-b50b-d073efb27f37', #Microsoft Application Virtualization 4.5
                #'c755e211-dc2b-45a7-be72-0bdc9015a63b', #Microsoft Application Virtualization 4.6
                #'1406b1b4-5441-408f-babc-9dcb5501f46f', #Microsoft Application Virtualization 5.0
                #'8f7c8263-d1eb-4144-89f6-fd568ec1364b', #Microsoft Azure Information Protection Client
                #'e903c733-c905-4b1c-a5c4-3528b6bbc746', #Microsoft Azure Site Recovery Provider
                #'7e903438-3690-4cf0-bc89-2fc34c26422b', #Microsoft BitLocker Administration and Monitoring v1
                #'2f3d1aba-2192-47b4-9c8d-87b41f693af4', #Microsoft Dynamics CRM 2011
                #'587f7961-187a-4419-8972-318be1c318af', #Microsoft Dynamics CRM 2011 SHS
                #'260d4ca6-768f-4e3e-9285-c30693bb7bc1', #Microsoft Dynamics CRM 2013
                #'3a78cd53-79b0-43a6-82f6-d9d6b9eec011', #Microsoft Dynamics CRM 2015
                #'734658e2-c499-46ac-953f-287b14deeb44', #Microsoft Dynamics CRM 2016
                #'25af568d-88b3-4cad-b694-07bc7f6adf24', #Microsoft Dynamics CRM 2016 SHS
                #'5e870422-bd8f-4fd2-96d3-9c5d9aafda22', #Microsoft Lync 2010
                #'04d85ac2-c29f-4414-9cb6-5bcd6c059070', #Microsoft Lync Server 2010
                #'01ce995b-6e10-404b-8511-08142e6b814e', #Microsoft Lync Server 2013
                #'f3869cc3-897b-4339-bb10-32ab2c765862', #Microsoft Monitoring Agent
                #'b0247430-6f8d-4409-b39b-30de02286c71', #Microsoft Online Services Sign-In Assistant
                #'a8f50393-2e42-43d1-aaf0-92bec8b60775', #Microsoft Research AutoCollage 2008
                #'e9ece729-676d-4b57-b4d1-7e0ab0589707', #Microsoft SQL Server 2008 R2 - PowerPivot for Microsoft Excel 2010
                #'56750722-19b4-4449-a547-5b68f19eee38', #Microsoft SQL Server 2012
                #'caab596c-64f2-4aa9-bbe3-784c6e2ccf9c', #Microsoft SQL Server 2014
                #'93f0b0bc-9c20-4ca5-b630-06eb4706a447', #Microsoft SQL Server 2016
                #'ca6616aa-6310-4c2d-a6bf-cae700b85e86', #Microsoft SQL Server 2017
                #'dee854fd-e9d2-43fd-bbc3-f7568e3ce324', #Microsoft SQL Server Management Studio v17
                #'a73eeffa-5729-48d4-8bf4-275132338629', #Microsoft StreamInsight V1.0
                #'bf6a6018-83f0-45a6-b9bf-074a78ec9c82', #Microsoft System Center DPM 2010
                #'29fd8922-db9e-4a97-aa00-ca980376b738', #Microsoft System Center Virtual Machine Manager 2007
                #'7e5d0309-78dd-4f52-a756-0259f88b634b', #Microsoft System Center Virtual Machine Manager 2008
                #'b790e43b-f4e4-48b4-9f0c-499194f00841', #Microsoft Works 8
                #'e9c87080-a759-475a-a8fa-55552c8cd3dc', #Microsoft Works 9
                #'6b9e8b26-8f50-44b9-94c6-7846084383ec', #MS Security Essentials
                #'4217668b-66f0-42a0-911e-a334a5e4dbad', #Network Monitor 3
                #'8508af86-b85e-450f-a518-3b6f8f204eea', #New Dictionaries for Microsoft IMEs
                #'6248b8b1-ffeb-dbd9-887a-2acf53b09dfe', #Office 2002/XP
                #'1403f223-a63f-f572-82ba-c92391218055', #Office 2003
                #'041e4f9f-3a3d-4f58-8b2f-5e6fe95c4591', #Office 2007
                #'84f5f325-30d7-41c4-81d1-87a0e6535b66', #Office 2010
                #'704a0a4a-518f-4d69-9e03-10ba44198bd5', #Office 2013
                #'25aed893-7c2d-4a31-ae22-28ff8ac150ed', #Office 2016
                #'6c5f2e66-7dbc-4c59-90a7-849c4c649d7a', #Office 2019
                #'30eb551c-6288-4716-9a78-f300ec36d72b', #Office 365 Client
                #'e164fc3d-96be-4811-8ad5-ebe692be33dd', #Office Communications Server 2007
                #'22bf57a8-4fe1-425f-bdaa-32b4f655284b', #Office Communications Server 2007 R2
                #'8bc19572-a4b6-4910-b70d-716fecffc1eb', #Office Communicator 2007 R2
                #'ec231084-85c2-4daf-bfc4-50bbe4022257', #Office Live Add-in
                #'d123907b-ba63-40cb-a954-9b8a4481dded', #OneCare Family Safety Installation
                #'dd78b8a1-0b20-45c1-add6-4da72e9364cf', #OOBE ZDP
                #'f0474daf-de38-4b6e-9ad6-74922f6f539d', #Photo Gallery Installation and Upgrades
                #'7cf56bdd-5b4e-4c04-a6a6-706a2199eff7', #Report Viewer 2005
                #'79adaa30-d83b-4d9c-8afd-e099cf34855f', #Report Viewer 2008
                #'f7f096c9-9293-422d-9be8-9f6e90c2e096', #Report Viewer 2010
                #'ce62f77a-28f3-4d4b-824f-0f9b53461d67', #Search Enhancement Pack
                #'6cf036b9-b546-4694-885a-938b93216b66', #Security Essentials
                #'50bb1d21-f01a-451c-9b1f-6c41e3c43ee7', #Service Bus for Windows Server 1.1
                #'9f3dd20a-1004-470e-ba65-3dc62d982958', #Silverlight
                #'8184d953-8366-4e13-8566-df0e15aca108', #Skype for Business Server 2015
                #'bec76be5-7aa9-497f-b70b-5fd1cfd1e3b1', #Skype for Business Server 2015, SmartSetup
                #'7145181b-9556-4b11-b659-0162fa9df11f', #SQL Server 2000
                #'60916385-7546-4e9b-836e-79d65e517bab', #SQL Server 2005
                #'c5f0b23c-e990-4b71-9808-718d353f533a', #SQL Server 2008
                #'bb7bc3a7-857b-49d4-8879-b639cf5e8c3c', #SQL Server 2008 R2
                #'7fe4630a-0330-4b01-a5e6-a77c7ad34eb0', #SQL Server 2012 Product Updates for Setup
                #'892c0584-8b03-428f-9a74-224fcd6887c0', #SQL Server 2014-2016 Product Updates for Setup
                #'c96c35fc-a21f-481b-917c-10c4f64792cb', #SQL Server Feature Pack
                #'b2b5aff0-734b-44e7-9934-48467fcae134', #Surface Hub 2S drivers
                #'58de46e5-6ccb-4154-91c1-73f8f4b84ce8', #System Center 1801 - Orchestrator
                #'daa70353-99b4-4e04-b776-03973d54d20f', #System Center 2012 - App Controller
                #'b0c3b58d-1997-4b68-8d73-ab77f721d099', #System Center 2012 - Data Protection Manager
                #'bf05abfb-6388-4908-824e-01565b05e43a', #System Center 2012 - Operations Manager
                #'ab8df9b9-8bff-4999-aee5-6e4054ead976', #System Center 2012 - Orchestrator
                #'6ed4a93e-e443-4965-b666-5bc7149f793c', #System Center 2012 - Virtual Machine Manager
                #'e54f3c9b-eec3-48f4-a791-ef1e2b0586d0', #System Center 2012 R2 - Data Protection Manager
                #'2a9170d5-3434-4820-885c-61a4f3fc6f84', #System Center 2012 R2 - Operations Manager
                #'6ddf2e90-4b40-471c-a664-6cd6b7e0d0a7', #System Center 2012 R2 - Orchestrator
                #'8a3485af-4301-43e1-b2d9-f9ddb7576125', #System Center 2012 R2 - Virtual Machine Manager
                #'50d71efd-1e60-4898-9ef5-f31a77bde4b0', #System Center 2012 SP1 - App Controller
                #'dd6318d7-1cff-44ed-a0b1-9d410c196792', #System Center 2012 SP1 - Data Protection Manager
                #'80d30b43-f814-41fd-b7c5-85c91ea66c45', #System Center 2012 SP1 - Operation Manager
                #'ba649061-a2bd-42a9-b7c3-825ce12c3cd6', #System Center 2012 SP1 - Virtual Machine Manager
                #'bc48031f-9353-4db2-a305-541e324374e2', #System Center 2016 - Data Protection Manager
                #'cc4ab3ac-9d9a-4f53-97d3-e0d6de39d119', #System Center 2016 - Operations Manager
                #'e505a854-6941-484f-a107-ebf0d1b64820', #System Center 2016 - Orchestrator
                #'d06f861b-2952-4063-bad5-ae8212746a61', #System Center 2016 - Virtual Machine Manager
                #'ae4500e9-17b0-4a78-b088-5b056dbf452b', #System Center Advisor
                #'d22b3d16-bc75-418f-b648-e5f3d32490ee', #System Center Configuration Manager 2007
                #'24c467b8-2fec-4f6c-bf32-d8f623b00b37', #System Center Data Protection Manager
                #'5f21acd2-d667-44f9-8d5e-485433e4d25c', #System Center Operations Manager 1807
                #'bcc7f992-8328-4e5f-b7bb-50d9a77d2343', #System Center Version 1801 - Virtual Machine Manager
                #'5a456666-3ac5-4162-9f52-260885d6533a', #Systems Management Server 2003
                #'ae4483f4-f3ce-4956-ae80-93c18d8886a6', #Threat Management Gateway Definition Updates for Network Inspection System
                #'cd8d80fe-5b55-48f1-b37a-96535dca6ae7', #TMG Firewall Client
                #'c8a4436c-1043-4288-a065-0f37e9640d60', #Virtual PC
                #'f61ce0bd-ba78-4399-bb1c-098da328f2cc', #Virtual Server
                #'a0dd7e72-90ec-41e3-b370-c86a245cd44f', #Visual Studio 2005
                #'e3fde9f8-14d6-4b5c-911c-fba9e0fc9887', #Visual Studio 2008
                #'c9834186-a976-472b-8384-6bb8f2aa43d9', #Visual Studio 2010
                #'cbfd1e71-9d9e-457e-a8c5-500c47cfe9f3', #Visual Studio 2010 Tools for Office Runtime
                #'e1c753f2-9f79-4577-b75b-913f4230feee', #Visual Studio 2010 Tools for Office Runtime
                #'abddd523-04f4-4f8e-b76f-a6c84286cc67', #Visual Studio 2012
                #'cf4aa0fc-119d-4408-bcba-181abb69ed33', #Visual Studio 2013
                #'1731f839-8830-4b9c-986e-82ee30b24120', #Visual Studio 2015
                #'a3c2375d-0c8a-42f9-bce0-28333e198407', #Windows 10
                #'05eebf61-148b-43cf-80da-1c99ab0b8699', #Windows 10 and later drivers
                #'34f268b4-7e2d-40e1-8966-8bb6ea3dad27', #Windows 10 and later upgrade & servicing drivers
                #'bab879a4-c1af-4b52-9617-0f9ae1286fb6', #Windows 10 Anniversary Update and Later Servicing Drivers
                #'0ba562e6-a6ba-490d-bdce-93a770ba8d21', #Windows 10 Anniversary Update and Later Upgrade & Servicing Drivers
                #'cfe7182c-14a0-4d7e-9f5e-505d5c3a66f6', #Windows 10 Creators Update and Later Servicing Drivers
                #'f5b5092c-d05e-4eb1-8a6a-919770378ff6', #Windows 10 Creators Update and Later Servicing Drivers
                #'06da2f0c-7937-4e28-b46c-a37317eade73', #Windows 10 Creators Update and Later Upgrade & Servicing Drivers
                #'e4b04398-adbd-4b69-93b9-477322331cd3', #Windows 10 Dynamic Update
                #'876ad18f-f41d-442a-ac64-f5c5ce74cc83', #Windows 10 Fall Creators Update and Later Servicing Drivers
                #'c70f1038-66ac-443d-9e58-ac22e891e4fb', #Windows 10 Fall Creators Update and Later Upgrade & Servicing Drivers
                #'e104dd76-2895-41c4-9eb5-c483a61e9427', #Windows 10 Feature On Demand
                #'abc45868-0c9c-4bc0-a36d-03d54113baf4', #Windows 10 GDR-DU
                #'3efabf46-3037-4c85-a752-3189e574b621', #Windows 10 GDR-DU FOD
                #'6111a83d-7a6b-4a2c-a7c2-f222eebcabf4', #Windows 10 GDR-DU LP
                #'7d247b99-caa2-45e4-9c8f-6d60d0aae35c', #Windows 10 Language Interface Packs
                #'fc7c9913-7a1e-4b30-b602-3c62fffd9b1a', #Windows 10 Language Packs
                #'d2085b71-5f1f-43a9-880d-ed159016d5c6', #Windows 10 LTSB
                #'c1006636-eab4-4b0b-b1b0-d50282c0377e', #Windows 10 S and Later Servicing Drivers
                #'bb06ba08-3df8-4221-8794-18effb79156a', #Windows 10 S Version 1709 and Later Servicing Drivers for testing
                #'b7f52cfb-c9e9-4481-9bc0-c8b4e208ba39', #Windows 10 S Version 1709 and Later Upgrade & Servicing Drivers for testing
                #'e727f134-a089-4b23-83f1-3004e054f658', #Windows 10 S Version 1803 and Later Servicing Drivers
                #'761370fd-6dbb-427f-899e-c19d56e22a9b', #Windows 10 S Version 1803 and Later Upgrade & Servicing Drivers
                #'39d54f77-4f1f-4e46-9752-c2de4cf2244d', #Windows 10 S, version 1809 and later, Servicing Drivers
                #'bb0dab86-78bd-4561-a71c-fb0071efd262', #Windows 10 S, version 1809 and later, Upgrade & Servicing Drivers
                #'8570b1a2-0551-42c8-a3e7-d3783c3d36d4', #Windows 10 version 1803 and Later Servicing Drivers
                #'29e060d2-aa33-4784-9b50-2021bb84cc18', #Windows 10 Version 1803 and Later Upgrade   & Servicing Drivers
                #'7ddc06c4-f2ff-4bb0-bc87-17b385c89a63', #Windows 10, version 1809 and later, Servicing Drivers
                #'13610e13-fac1-4017-b703-82062db96be4', #Windows 10, version 1809 and later, Upgrade & Servicing Drivers
                #'b3c75dc1-155f-4be4-b015-3f1a91758e52', #Windows 10, version 1903 and later
                #'3b4b8621-726e-43a6-b43b-37d07ec7019f', #Windows 2000
                #'bfe5b177-a086-47a0-b102-097e4fa1f807', #Windows 7
                #'2ee2ad83-828c-4405-9479-544d767993fc', #Windows 8
                #'393789f5-61c1-4881-b5e7-c47bcca90f94', #Windows 8 Dynamic Update
                #'589db546-7849-47f5-bbc0-1f66cf12f5c2', #Windows 8 Embedded
                #'3e5cc385-f312-4fff-bd5e-b88dcf29b476', #Windows 8 Language Interface Packs
                #'97c4cee8-b2ae-4c43-a5ee-08367dab8796', #Windows 8 Language Packs
                #'6407468e-edc7-4ecd-8c32-521f64cee65e', #Windows 8.1
                #'405706ed-f1d7-47ea-91e1-eb8860039715', #Windows 8.1 and later drivers
                #'f7b29b7a-086b-43f9-9cc8-e1a2f8a31e08', #Windows 8.1 Drivers
                #'18e5ea77-e3d1-43b6-a0a8-fa3dbcd42e93', #Windows 8.1 Dynamic Update
                #'14a011c7-d17b-4b71-a2a4-051807f4f4c6', #Windows 8.1 Language Interface Packs
                #'01030579-66d2-446e-8c65-538df07e0e44', #Windows 8.1 Language Packs
                #'b1b8f641-1ff2-4ae6-b247-4fe7503787be', #Windows Admin Center
                #'1aea70f3-d989-4f89-9055-b0bc9945b75f', #Windows Azure Pack: Admin API
                #'983dabe5-e68d-4cb3-ae5e-6da88e66783f', #Windows Azure Pack: Admin Authentication Site
                #'2f1d3c10-1e92-487b-baba-2c1c645367b9', #Windows Azure Pack: Admin Site
                #'57869cb9-cd47-4ce4-acd5-caf49a0c713f', #Windows Azure Pack: Configuration Site
                #'6102ab07-dd96-4407-8c82-2f2db7022248', #Windows Azure Pack: Microsoft Best Practice Analyzer
                #'10b00347-cd06-41fd-b7ba-32200693e114', #Windows Azure Pack: Monitoring Extension
                #'71debf20-7fce-4e93-8a6c-4a3fad0313ec', #Windows Azure Pack: MySQL Extension
                #'19243b1e-a4c1-4e87-80f4-fa8546ce4489', #Windows Azure Pack: PowerShell API
                #'8516af00-35dc-4fd6-af4f-e1a9f117a882', #Windows Azure Pack: SQL Server Extension
                #'2c25d763-d623-433f-b956-0de582e32b19', #Windows Azure Pack: Tenant API
                #'45afcceb-93c4-4ac3-909c-ca349acbc264', #Windows Azure Pack: Tenant Authentication Site
                #'3f50dcc0-6199-4ae0-a166-6d87d4e6f83e', #Windows Azure Pack: Tenant Public API
                #'95a5f8e0-f2ab-4be6-bc4a-34d4b790192f', #Windows Azure Pack: Tenant Site
                #'9e185861-6465-41db-83c4-bb1480a55851', #Windows Azure Pack: Usage Extension
                #'5c91542d-b573-44e9-a86d-b13b27cd98db', #Windows Azure Pack: Web App Gallery Extension
                #'3c9e83e3-614d-4670-9205-cfcf3ea62a29', #Windows Azure Pack: Web Sites
                #'8c3fcc84-7410-4a95-8b89-a166a0190486', #Windows Defender
                #'50c04525-9b15-4f7c-bed4-87455bcd7ded', #Windows Dictionary Updates
                #'f14be400-6024-429b-9459-c438db2978d4', #Windows Embedded Developer Update
                #'f4b9c883-f4db-4fb5-b204-3343c11fa021', #Windows Embedded Standard 7
                #'e9b56b9a-0ca9-4b3e-91d4-bdcf1ac7d94d', #Windows Essential Business Server 2008
                #'6966a762-0c7c-4261-bd07-fb12b4673347', #Windows Essential Business Server 2008 Setup Updates
                #'649f3e94-ed2f-42e8-a4cd-e81489af357c', #Windows Essential Business Server Preinstallation Tools
                #'470bd53a-c36a-448f-b620-91feede01946', #Windows GDR-Dynamic Update
                #'cb263e3f-6c5a-4b71-88fa-1706f9549f51', #Windows Internet Explorer 7 Dynamic Installer
                #'5312e4f1-6372-442d-aeb2-15f2132c9bd7', #Windows Internet Explorer 8 Dynamic Installer
                #'0ea196ba-7a32-4e76-afd8-46bd54ecd3c6', #Windows Live
                #'afd77d9e-f05a-431c-889a-34c23c9f9af5', #Windows Live
                #'5ea45628-0257-499b-9c23-a6988fc5ea85', #Windows Live Toolbar
                #'e88a19fb-a847-4e3d-9ae2-13c2b84f58a6', #Windows Media Dynamic Installer
                #'8c27cdba-6a1c-455e-af20-46b7771bbb96', #Windows Next Graphics Driver Dynamic update
                #'0a07aea1-9d09-4c1e-8dc7-7469228d8195', #Windows RT
                #'2c62603e-7a60-4832-9a14-cfdfd2d71b9a', #Windows RT 8.1
                #'dd1aa213-54e7-4173-8456-b278964a26b6', #Windows Safe OS Dynamic Update
                #'dbf57a08-0d5a-46ff-b30c-7715eb9498e9', #Windows Server 2003
                #'7f44c2a7-bc36-470b-be3b-c01b6dc5dd4e', #Windows Server 2003, Datacenter Edition
                #'ba0ae9cc-5f01-40b4-ac3f-50192b5d6aaf', #Windows Server 2008
                #'fdfe8200-9d98-44ba-a12a-772282bf60ef', #Windows Server 2008 R2
                #'ec9aaca2-f868-4f06-b201-fb8eefd84cef', #Windows Server 2008 Server Manager Dynamic Installer
                #'a105a108-7c9b-4518-bbbe-73f0fe30012b', #Windows Server 2012
                #'26cbba0f-45de-40d5-b94a-3cbe5b761c9d', #Windows Server 2012 Language Packs
                #'d31bd4c3-d872-41c9-a2e7-231f372588cb', #Windows Server 2012 R2
                #'f3c2263d-b256-4c49-a246-973c0e366449', #Windows Server 2012 R2  and later drivers
                #'bfd3e48c-c96b-43fd-8b09-98cdc89dc77e', #Windows Server 2012 R2 Drivers
                #'8b4e84f6-595f-41ed-854f-4ca886e317a5', #Windows Server 2012 R2 Language Packs
                #'569e8e8f-c6cd-42c8-92a3-efbb20a0f6f5', #Windows Server 2016
                #'3c54bb6c-66d1-4a79-884c-8a0c96fa20d1', #Windows Server 2016 and Later Servicing Drivers
                #'f702a48c-919b-45d6-9aef-ca4248d50397', #Windows Server 2019
                #'b645dd0f-2965-43e8-b055-8ea47e2d71d7', #Windows Server 2019 and later, Servicing Drivers
                #'ecf560de-38d7-4aa0-beef-e74041c581a4', #Windows Server 2019 and later, Upgrade & Servicing Drivers
                #'323cceaf-b60b-4a0d-8a8a-3069efde76bf', #Windows Server Drivers
                #'4e487029-f550-4c22-8b31-9173f3f95786', #Windows Server Manager – Windows Server Update Services (WSUS) Dynamic Installer
                #'eef074e9-61d6-4dac-b102-3dbe15fff3ea', #Windows Server Solutions Best Practices Analyzer 1.0
                #'0b378f2d-bff3-47dd-9b7f-5c9f966bdd81', #Windows Server Technical Preview Language Packs
                #'21210d67-50bc-4254-a695-281765e10665', #Windows Server, version 1903 and later
                #'032e3af5-1ac5-4205-9ae5-461b4e8cd26d', #Windows Small Business Server 2003
                #'575d68e2-7c94-48f9-a04f-4b68555d972d', #Windows Small Business Server 2008
                #'7fff3336-2479-4623-a697-bcefcf1b9f92', #Windows Small Business Server 2008 Migration Preparation Tool
                #'1556fc1d-f20e-4790-848e-90b7cdbedfda', #Windows Small Business Server 2011 Standard
                #'e7441a84-4561-465f-9e0e-7fc16fa25ea7', #Windows Ultimate Extras
                #'26997d30-08ce-4f25-b2de-699c36a8033a', #Windows Vista
                #'90e135fb-ef48-4ad0-afb5-10c4ceb4ed16', #Windows Vista Dynamic Installer
                #'a901c1bd-989c-45c6-8da0-8dde8dbb69e0', #Windows Vista Ultimate Language Packs
                #'558f4bc3-4827-49e1-accf-ea79fd72d4c9', #Windows XP
                #'a4bedb1d-a809-4f63-9b49-3fe31967b6d0', #Windows XP 64-Bit Edition Version 2003
                #'874a7757-3a13-43b2-a7f2-cf2ff43dd6bf', #Windows XP Embedded
                #'4cb6ebd5-e38a-4826-9f76-1416a6f563b0', #Windows XP x64 Edition
                #'81b8c03b-9743-44b1-8c78-25e750921e36', #Works 6-9 Converter
                #'a13d331b-ce8f-40e4-8a18-227bf18f22f3', #Writer Installation and Upgrades
            }
            $updateClassification = {
                #'051f8713-e600-4bee-b7b7-690d43c78948', #WSUS Infrastructure Updates
                #'0fa1201d-4330-4fa8-8ae9-b877473b6441', #Security Updates
                #'28bc880e-0592-4cbf-8f95-c79b17911d5f', #Update Rollups
                #'3689bdc8-b205-4af4-8d4a-a63924c5e9d5', #Upgrades
                #'5c9376ab-8ce6-464a-b136-22113dd69801', #Applications
                #'68c5b0a3-d1a6-4553-ae49-01d3a7827828', #Service Packs
                #'77835c8d-62a7-41f5-82ad-f28d1af1e3b1', #Driver Sets
                #'b4832bd8-e735-4761-8daf-37f882276dab', #Tools
                #'b54e7d24-7add-428f-8b75-90a396fa584f', #Feature Packs
                #'cd5ffd1e-e932-4e3a-bf74-18bf0b1bbd83', #Updates
                #'e0789628-ce08-4437-be74-2495b842f43b', #Definition Updates
                #'e6cf1350-c01b-414d-a61f-263d14d133b4', #Critical Updates
                #'ebfc1fc5-71a4-4f7b-9aca-3b9a503104a0', #Drivers
            }
        }
    }      
#endregion 
# --------------------------------------------------------------------------------------------
#region MAIN_SUB

Function MainSub{
# ===============================================================================================================================================================================
#region 1_PRE-CHECKS            
    Write-Log -iTabs 1 "Starting 1 - Pre-Checks." -bConsole $true -scolor Cyan
    #region 1.0 Checking/Loading SCCM Powershell Module                
        Write-Log -iTabs 2 "1.0 Checking/Loading SCCM Powershell module from $($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')" -bConsole $true -scolor Cyan
        if (Get-Module | Where-Object {$_.Name -like "*ConfigurationManager*"}){            
            Write-Log -iTabs 3 "SCCM PS Module was found loaded in this session!" -bConsole $true -scolor Green
        }
        else{            
            Write-Log -iTabs 3 "SCCM PS Module was not found in this session! Loading Module. This might take a few minutes..." -bConsole $true
            Try{                            
                Write-Log  -iTabs 4 "Looking for Module in $(($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1'))" -bConsole $true
                Import-module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')                
                Write-Log  -iTabs 4 "Successfully loaded SCCM Powershell module"  -bConsole $true -scolor Green
            }
            catch{                
                Write-Log -iTabs 4 "Unable to Load SCCM Powershell module." -bConsole $true -scolor Red                
                Write-Log -iTabs 4 "Aborting script." -iTabs 4 -bConsole $true -scolor Red
                $global:iExitCode = 9001
                return $global:iExitCode
            }                      
        }
    #endregion    
    #region 1.1 Confirm Script Arguments            
        Write-Log -iTabs 2 "1.1 Checking Script arguments" -bConsole $TRUE -sColor Cyan
        Write-Log -iTabs 3 "Script is running with Command Line: $sCMDArgs" -bConsole $true -bTxtLog $false        
        if (($Action -eq "Auto-Run") -and ($Scope -eq "Other")){                
            Write-Log -itabs 3 "Action 'Auto-Run' is not supported with Scope 'Other'." -bConsole $true
            Write-Log -itabs 4 "----------------------------------------------" -bConsole $true -sColor red
            HowTo-Script                
            Write-Log -itabs 4 "----------------------------------------------" -bConsole $true -sColor red
            Write-Log -itabs 4 "Aborting script." -bConsole $true -sColor red
            $global:iExitCode = 9002
            return $global:iExitCode
        }
        if ($Scope -eq "Other"){                        
            Write-Log -iTabs 3 "Scope 'Other' requires data to be collected from User." -bConsole $true
            #Setting SMS Provider
            if ($null -eq $SMSProvider){                                
                $smsProvTest = $false
                do{
                    $SMSProvider = Read-Host "                                      SMS Provider [<ServerFQDN>/Abort] "                    
                    if ($SMSProvider -eq "Abort"){                        
                        Write-Log -iTabs 5 "Aborting script." -bConsole $true -sColor red
                        $global:iExitCode = 5001
                        return $global:iExitCode
                    }
                    Write-Log -iTabs 5 "User set '$SMSProvider' as SMSProvider"
                    Write-Log -iTabs 5 "Testing '$SMSProvider' connection..." -bConsole $true
                    if (Test-Connection -ComputerName $SMSProvider -Count 1 -Quiet){
                        Write-Log -iTabs 5 "$SMSProvider was found and set as SMSProvider" -bConsole $true -sColor green                        
                        $smsProvTest = $true
                    }
                    else{
                        Write-Log -iTabs 5 "Unable to reach $SMSProvider. Ensure server FQDN is valid" -bConsole $true -sColor red                                                
                        $smsProvTest = $false
                    }                
                }while(!$smsProvTest)                
            }  
            #Setting SCCM Site        
            if ($null -eq $SCCMSite){
                $sccmSiteTest = $false
                do{
                    $SCCMSite = Read-Host "                                      SCCM Site [<SITECODE>/Abort] "
                    if ($SCCMSite -eq "Abort"){
                        Write-Log -iTabs 5 "Aborting script." -bConsole $true -sColor red
                        $global:iExitCode = 5001
                        return $global:iExitCode
                    }
                    Write-Log -iTabs 5 "User set '$SCCMSite' as SCCM Site..."
                    Write-Log -iTabs 5 "Testing '$SCCMSite' as SCCM Site..." -bConsole $true
                    try{
                        $qrySccmSite = $(get-WMIObject -ComputerName $SMSProvider -Namespace "root\SMS" -Class "SMS_ProviderLocation" | Where-Object {$_.ProviderForLocalSite -eq "True"} | Select-Object Sitecode).Sitecode
                    }
                    catch{
                        Write-Log -iTabs 5 "Unable to reach $SMSProvider SiteCode via WMI. Ensure user permissions are present for this operation." -bConsole $true -sColor red
                        $sccmSiteTest=$false
                    }
                    if ($qrySccmSite -eq $SCCMSite){
                        Write-Log -iTabs 5 "SCCM Site $SCCMSite found in $SMSProvider. Setting as SCCM Site Code" -bConsole $true -sColor green                        
                        $sccmSiteTest=$true
                    }
                    else{
                        Write-Log -iTabs 5 "SCCM Site $SCCMSite not found in $SMSProvider. Verify Site is valid" -bConsole $true -sColor red                        
                        $sccmSiteTest=$false
                    }                
                }while(!$sccmSiteTest)                  
            }
            #Setting Template Name
            if ($null -eq $TemplateName){
                $sugTest = $false
                do{
                    $TemplateName = Read-Host "                                      Template Name [<SUGName>/Abort] "
                    if ($TemplateName -eq "Abort"){
                        Write-Log -iTabs 5 "Aborting script." -bConsole $true -sColor red
                        $global:iExitCode = 5001
                        return $global:iExitCode
                    }       
                    else{       
                        Write-Log -iTabs 5 "Template Name was set as '$TemplateName'" -bConsole $true
                        $answer = Read-Host "                                          Do you confirm? [Y/n] "                
                        if ($answer -eq "Y"){
                            Write-Log -iTabs 6 "User confirmed Template Name"
                            $sugTest=$true
                        }
                        else{
                            Write-Log -iTabs 6 "User cleared Template Name"
                        }
                    }      
                }while(!$sugTest)                            
            }
        }
        # Testing SCCM Drive
        try{
            $originalLocation = Get-Location
            Write-Log -iTabs 3 "Connecting to SCCM PS Location at $($SCCMSite):\" -bConsole $true -bTxtLog $false        
            Set-Location $SCCMSite":"            
            Write-Log -iTabs 4 "Connected to SCCM PS Location at $($SCCMSite):\" -bConsole $true -sColor Green      
        }
        catch{
            Write-Log -iTabs 4 "Unable to connect to SCCM PSDrive. Aborting Script" -bConsole $true -sColor red
            Write-Log -iTabs 4 "Aborting script." -bConsole $true -sColor red
            $global:iExitCode = 9003            
            return $global:iExitCode
        }            
        #Confirming Settings                              
        Write-Log -iTabs 3 "Setings were defined as:" -bConsole $true        
        Write-Log -iTabs 4 "SCCM Scope:                  $Scope" -bConsole $true -sColor Yellow        
        Write-Log -iTabs 4 "SMSProvider:                 $SMSProvider" -bConsole $true                
        Write-Log -iTabs 4 "SCCM Site Code:              $SCCMSite" -bConsole $true                 
        Write-Log -iTabs 4 "Update Age to be Reported:   $timeReport" -bConsole $true
        Write-Log -iTabs 4 "Severity:                    $severity" -bConsole $true
        Write-Log -iTabs 4 "Products to be Considered:   $updateProducts" -bConsole $true
        Write-Log -iTabs 4 "Classifications:             $updateClassification" -bConsole $true

    #endregion  
    #region 1.2 Is this SCCM Admin User?            
        Write-Log -iTabs 2 "1.2 Checking if user has permissions in SCCM to run this script..." -bConsole $true -sColor Cyan        
        <#    
        $userRoles = (Get-CMAdministrativeUser -Name $($sUserDomain+"\"+$sUserName)).RoleNames
        foreach ($role in $userRoles){
            If (($role -eq "Full Administrator") -or ($role -eq "Software Update Manager")){
                $userRoleTest = $true
            }
        }
        if ($userRoleTest){
            Write-Host  "        User has permissions to execute this script." -ForegroundColor Green            
            Write-Log           "User has permissions to execute this script." -iTabs 3
        }
        else{
            Write-Host "        User does not have permissions to execute this script." -ForegroundColor Red
            Write-Log          "User does not have permissions to execute this script." -iTabs 3
            Write-Host "        Aborting script." -ForegroundColor Red
            Write-Log          "Aborting script." -iTabs 3
            $global:iExitCode = 9002
            return $global:iExitCode
        }    
        #>                    
        Write-Log  -iTabs 3 "Pre-Check to be implemented" -bConsole $true
    #endregion    
    #region 1.3 Querying Software Update Information        
        Write-Log -iTabs 2 "1.3 Getting Software Update Process Information" -bConsole $true -sColor Cyan
            #region Checking ADR                
                Write-Log -iTabs 3 "Checking if ADR Named ($($TemplateName)ADR) is present." -bConsole $true
                try{
                    $defaultAdr = Get-CMAutoDeploymentRule -fast -Name "$($TemplateName)ADR"
                    if ($defaultAdr.Count -gt 0){                        
                        Write-Log -iTabs 4 "ADR ($($TemplateName)ADR) was found in SCCM Environment." -bConsole $true
                    }
                    else{                        
                        Write-Log -iTabs 4 "ADR ($($TemplateName)ADR) was not found in SCCM Environment." -bConsole $true -sColor yellow
                        Write-Log -iTabs 4 "It is strongly recomended to have an ADR responsible for creating Monhtly updates." -bConsole $true -sColor yellow
                    }
                }
                catch{                    
                    Write-Log -iTabs 4 "Unable to verify existing ADRs. Permission Error. Ensure script is running with SCCM Full Admin permissionts and access to SCCM WMI Provider." -bConsole $true -sColor red                    
                    Write-Log -iTabs 4 "Aborting script." -bConsole $true -sColor red
                    $global:iExitCode = 9004
                    return $global:iExitCode
                }
            #endregion
            #region Checking SUGs           
                Write-Log -iTabs 3 "Checking if SUGs are present." -bConsole $true
                #region Gettings SUG Info         
                try{
                        $sugs = Get-CMSoftwareUpdateGroup | Where-Object {$_.LocalizedDisplayName -like "$TemplateName*"} | ConvertTo-Array                                               
                    }
                #Error while getting SUG Info
                catch{                                                                        
                        Write-Log -iTabs 4 "Unable to query Software Update Groups. Permission Error. Ensure script is running with SCCM Full Admin permissionts and access to SCCM WMI Provider." -bConsole $true -sColor red                        
                        Write-Log -iTabs 4 "Aborting script." -bConsole $true -sColor red
                        $global:iExitCode = 9005
                        return $global:iExitCode
                    }      
                #endregion             
                #region sugAged
                #sugAged was found
                if (($sugs | Where-Object {$_.LocalizedDisplayName -eq $TemplateName+"Aged"}).Count -gt 0){                        
                    Write-Log -iTabs 4 "$($TemplateName)Aged was found." -bConsole $true
                    $AgedSUG = $sugs | Where-Object {$_.LocalizedDisplayName -eq $TemplateName+"Aged"}
                }
                #sugAged was not found
                else{
                    Write-Log -iTabs 4 "$($TemplateName)Aged wasn't found. This SUG is required to proceed with script execution." -bConsole $true -sColor yellow
                    do{
                        if($action -like "*Run"){$answer = "Y"}
                        else{
                            $answer = Read-Host "                                      Do you want to create Software Update Group '$($TemplateName)Aged'? [Y/n] "                
                        }
                    } while (($answer -ne "Y") -and ($answer -ne "n"))
                    #aborting script
                    if ($answer -eq "n"){                
                        Write-Log -iTabs 5 "User don't want to create Software Update Group $($TemplateName)Aged at this moment"-bConsole $true -sColor red
                        Write-Log -iTabs 5 "Aborting script." -bConsole $true -sColor red
                        $global:iExitCode = 5002
                        return $global:iExitCode
                    }   
                    # Creating sugAged
                    if ($answer -eq "y"){                                            
                        Write-Log -iTabs 4 "Creating $($TemplateName)Aged..." -bConsole $true
                        try{
                            New-CMSoftwareUpdateGroup -Name "$($TemplateName)Aged" | Out-Null                                
                            Write-Log -iTabs 4 "$($TemplateName)Aged was created" -bConsole $true -sColor green                                                                
                        }    
                        catch{                                
                            Write-Log -iTabs 4 "Error while creating $($TemplateName)Aged. Ensure script is running with SCCM Full Admin permissionts and access to SCCM WMI Provider." -bConsole $true -sColor red
                            Write-Log -iTabs 4 "Aborting script." -bConsole $true -sColor red
                            $global:iExitCode = 9006
                            return $global:iExitCode                            
                        }
                        Write-Log -iTabs 4 "Reloading SUG Array." -bConsole $true      
                        $sugs = Get-CMSoftwareUpdateGroup | Where-Object {$_.LocalizedDisplayName -like "$TemplateName*"} | Sort LocalizedDisplayName | ConvertTo-Array                                            
                        $AgedSUG = $sugs | Where-Object {$_.LocalizedDisplayName -eq $TemplateName+"Aged"} #All Aged, but valid Updates                     
                    } 
                }
                #endregion                    
            #endregion           
            #region Query all Expired Updates            
            Write-Log -iTabs 3 "Getting all Expired KBs from SCCM WMI." -bConsole $true
            try{
                $ExpiredUpdates = Get-CMSoftwareUpdate -IsExpired $true -fast | Select-Object -Property CI_ID
                Write-Log -iTabs 4 "Expired KBs: $($ExpiredUpdates.Count)" -bConsole $true
            }
            catch{
                Write-Log -iTabs 4 "Error getting Update info from SCCM WMI." -bConsole $true -sColor red
                Write-Log -iTabs 4 "Aborting script."  -bConsole $true -sColor red
                $global:iExitCode = 9009
                return $global:iExitCode
            }
            #endregion
            #region Query All Superseded Updates
            Write-Log -iTabs 3 "Getting all Superseded KBs from SCCM WMI." -bConsole $true
            try{
                $SupersededUpdates = Get-CMSoftwareUpdate -IsSuperseded $true -fast | Select-Object -Property CI_ID
                Write-Log -iTabs 4 "Superseded KBs: $($SupersededUpdates.Count)" -bConsole $true
            }
            catch{
                Write-Log -iTabs 4 "Error getting Update info from SCCM WMI." -bConsole $true -sColor red
                Write-Log -iTabs 4 "Aborting script."  -bConsole $true -sColor red
                $global:iExitCode = 9009
                return $global:iExitCode
            }
            #endregion
            #region Query All Aged Updates            
            Write-Log -iTabs 3 "Getting all Aged KBs from SCCM WMI." -bConsole $true
            try{
                $AgedUpdates = Get-CMSoftwareUpdate -DatePostedMax $(Get-Date).AddDays(-$timeSustainerAge) -IsSuperseded $false -IsExpired $false -fast | Select-Object -Property CI_ID
                Write-Log -iTabs 4 "Aged KBs: $($AgedUpdates.Count)" -bConsole $true
            }
            catch{
                Write-Log -iTabs 4 "Error getting Update info from SCCM WMI." -bConsole $true -sColor red
                Write-Log -iTabs 4 "Aborting script."  -bConsole $true -sColor red
                $global:iExitCode = 9009
                return $global:iExitCode
            }
            #endregion
    #endregion    
    #region 1.4 Finalizing Pre-Checks      
    Write-Log -iTabs 2 "1.4 - Finalizing Pre-Checks:" -bConsole $true -sColor cyan    
    Write-Log -itabs 3 "SUG Information - These SUGs will be evaluated/changed by this script." -bConsole $true
    #$sugs | ft   
    $initNumUpdates=0 
    $initNumSugs=0
    foreach ($sug in $sugs | where {($sugs.LocalizedDisplayName -like $TemplateName -and $sugs.IsDeployed -eq $true) -or $sugs.LocalizedDisplayName -eq $TemplateName+"Aged"} | Sort-Object $sugs.LocalizedDisplayName){
        #$sugName = $sug.LocalizedDisplayName
        Write-Log -itabs 4 $sug.LocalizedDisplayName -bConsole $true
        $initNumUpdates+=$($sug.Updates).Count
        $initNumSugs+=1
    }    
    Write-Log -itabs 3 "Package Information - These PKGs will be evaluated/changed by this script." -bConsole $true            
    $initPkgSize=0
    $initPkgSize+=$pkgMonth.PackageSize/1024/1024
    Write-Log -itabs 4 "$($pkgMonth.PackageID) - $($pkgMonth.Name) - $([math]::Round($pkgMonth.PackageSize/1024/1024,2)) GB." -bConsole $true    
    Write-Log -itabs 4 "$($pkgAged.PackageID) - $($pkgAged.Name) - $([math]::Round($pkgAged.PackageSize/1024/1024,2)) GB." -bConsole $true
    $initPkgSize+=$pkgAged.PackageSize/1024/1024
            
    #$initNumUpdates = ($sugs | Where-Object {$_.LocalizedDisplayName -ne $SUGTemplateName+"Report" -and $_.LocalizedDisplayName -ne $SUGTemplateName+"Tracked" -and $_.LocalizedDisplayName -ne $SUGTemplateName+"MSFT"} | Measure-Object -Property NumberofUpdates -Sum).Sum
    Write-Log -itabs 3 "Number of Updates: $initNumUpdates" -bConsole $true
    #$initNumSugs = $sugs.Count
    Write-Log -itabs 3 "Number of SUGs: $initNumSugs" -bConsole $true    
    #$initPkgSize = ($pkgs | Measure-Object -Property PackageSize -Sum).Sum/1024/1024
    Write-Log -itabs 3 "Space used by Packages: $([math]::Round($initPkgSize,2)) GB" -bConsole $true    
    Write-Log -itabs 2 "Pre-Checks are complete. Script will make environment changes in the next interaction." -bConsole $true
    Write-Log -itabs 2 "Getting User confirmation to proceed"
    do{
        Write-Log -iTabs 3 "Above you have the list of Packages and Software Update Groups which will be managed by this script." -bConsole $true
        Write-Log -iTabs 3 "Review the list above and make sureare indeed the right targets for actions." -bConsole $true
        if ($action -eq "Check"){
            Write-Log -iTabs 3 "Script would make environment changes in the next interaction. Action Check identified. Script will log but not execute actions." -bConsole $true        
        }
        else{
            Write-Log -iTabs 3 "Script will make environment changes in the next interaction" -bConsole $true -scolor Yellow        
        }
        if($action -like "*run"){
            $answer = "Y"
        }
        else{
            $answer = Read-Host "                                  |Do you want to proceed? [Y/n]"     
        }   
    } while (($answer -ne "Y") -and ($answer -ne "n"))
    if ($answer -eq "n"){
        Write-Log -iTabs 3 "User Aborting script." -bConsole $true -sColor red
        $global:iExitCode = 8001
        return $global:iExitCode
    }
    else{
        Write-Log -iTabs 2 "User confirmation received." 
    }
    #endregion
    Write-Log -iTabs 1 "Completed 1 - Pre-Checks." -bConsole $true -sColor Cyan    
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================

# ===============================================================================================================================================================================
#region 2_EXECUTION
    Write-Log -iTabs 1 "Starting 2 - Execution."   -bConsole $true -sColor cyan    
   
    Write-Log -iTabs 1 "Completed 2 - Execution." -bConsole $true -sColor cyan
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================
        
# ===============================================================================================================================================================================
#region 3_POST-CHECKS
# ===============================================================================================================================================================================
    Write-Log -iTabs 1 "Starting 3 - Post-Checks." -bConsole $true -sColor cyan
    
    Write-Log -iTabs 1 "Completed 3 - Post-Checks." -bConsole $true -sColor cyan
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