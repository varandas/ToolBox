#REQUIRES -Version 4.0
<#
.Synopsis
   
.DESCRIPTION
   
.EXAMPLE
  
.EXAMPLE
  
.EXAMPLE
  
#>
param( 
    [string]$Path=".\wsusscn2.cab",
    [ValidateSet("RemoteLog","SQLLog","Both")][string]$LogType
)
# --------------------------------------------------------------------------------------------
#region HEADER
$SCRIPT_TITLE = "Get-MissingUpdates"
$SCRIPT_VERSION = "1.0"

$ErrorActionPreference 	= "Continue"	# SilentlyContinue / Stop / Continue

# -Script Name: Get-MissingUpdates.ps1------------------------------------------------------ 
# Based on PS Template Script Version: 1.0
# Author: Jose Varandas

#
# Owned By: Jose Varandas
# Purpose: Query machine for missing updates using external info source (Microsoft WSUSSCN2.cab). 
# 
#
#
# Dependencies: 
#                ID running script must be Local administrator
#                wsusscn2.cab must be present in the same folder or a valid path must be given via parameter
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
            Write-Log -sMessage "-Path -> Defines location for wsuscan2.cab. Default is `".\`"" -iTabs 3                        
            Write-Log -sMessage "-LogType (RemoteTxt/SQLLog/Both/None) -> Defines type of logging for missing updates." -iTabs 3        
                Write-Log -sMessage "-> None(default): Log is created locally only." -iTabs 4                        
                Write-Log -sMessage "-> RemoteTxt: Log is created locally and in a remote location. Share path will be required" -iTabs 4                        
                    Write-Log -sMessage "-> SharePath: Network Share which allows Domain Users and Domain Computers write access" -iTabs 5
                Write-Log -sMessage "-> SQLLog: Log is created locally and in a SQL Database location. SQL Server Name, SQL Instance and permissions to create/insert tables are required." -iTabs 4
                    Write-Log -sMessage "-> SQLServerName" -iTabs 5
                    Write-Log -sMessage "-> SQLInstanceName" -iTabs 5
                Write-Log -sMessage "-> Both: Log is created locally, in a remote location and in a SQL Database. Info will be required." -iTabs 4                   
                    Write-Log -sMessage "-> SharePath: Network Share which allows Domain Users and Domain Computers write access" -iTabs 5
                    Write-Log -sMessage "-> SQLServer" -iTabs 5
                    Write-Log -sMessage "-> SQLInstance" -iTabs 5
                    Write-Log -sMessage "-> SQLDB" -iTabs 5
    Write-Log -sMessage "============================================================================================================================" -iTabs 1            
    Write-Log -sMessage "EXAMPLE:" -iTabs 1
        Write-Log -sMessage ".\$sScriptName -Path `"C:\Users\Admin\Desktop\wsusscn2.cab`"" -iTabs 2     
            Write-Log -sMessage "Script will check for missing updates in WSUSScan CAB, SCCM WMI and log information locally" -iTabs 2     
        Write-Log -sMessage ".\$sScriptName -LogType Both -SQLServer SQL01 -SQLInstance MSSQLSERVER -SQLDB Any_DB ." -iTabs 2     
            Write-Log -sMessage "Script will check for missing updates in WSUSScan CAB, SCCM WMI. It will log information locally, remotely and in SQL Server" -iTabs 2  
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
            9001 - Error querying CCM_UpdatesStatus WMI class
            9002 - Error querying CCM_UpdateCIAssignment WMI class
            9999 - Unhandled Exception     

   
 Revision History: (Date, Author, Version, Changelog)
		2019/04/08 - Jose Varandas - 1.0			
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
    $sOutFileName	= "Get-MissingUpdates.log" # Log File Name    
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
    # $RemoteLogPath = "\\sccm01\Logs\SCCM\WSUS" #PVA
    $RemoteLogPath = "\\sccm01\Share\Logs" #VAR
    $RemoteLogName = "MissingUpdates$($cabver).log"
    $SQLServer = "SQL01"
    $SQLInstance = "MSSQLSERVER"
    $SQLDB = "MissingUpdates"
    $SQLTable = "MissingUpdates1904"
    if ($Path -eq ".\wsusscn2.cab"){
        $Path = Join-Path -Path $sScriptPath -ChildPath "wsusscn2.cab"
    }
    $cabVer="1903"
    # ****************************************************  
#endregion 
# --------------------------------------------------------------------------------------------
#region MAIN_SUB

Function MainSub{
# ===============================================================================================================================================================================
#region 1_PRE-CHECKS            
    Write-Log -iTabs 1 "Starting 1 - Pre-Checks."-scolor Cyan
    #region 1.1 Checking if Machine's SCCM WMI is reacheable
    Write-Log -iTabs 2 "Checking if machine has CCM_UpdateStatus class in WMI accessible."
    try{
        $sccmUpdateStatus = Get-WmiObject -Namespace ROOT\ccm\SoftwareUpdates\UpdatesStore -Class CCM_UpdateStatus -ErrorAction Stop
        Write-Log -iTabs 3 "WMI Class was found and update status loaded." -sColor Green
    }
    catch{
        Write-Log -iTabs 3 "WMI Class was not found. Script will proceed but won't be able to determine status for updates deployed by SCCM." -sColor Yellow
        $sccmUpdateStatus = $false
    }
    #endregion
    #region 1.2 Checking if WSUSSCN2.CAB is found
    Write-Log -iTabs 2 "Checking if wsusscn2.cab if found at $path."
    try{
        if (!(Test-Path $Path -ErrorAction Stop)){
            throw "CAB not found"
        }
        $wsusCabStatus = $true
        Write-Log -iTabs 3 "WSUSSCN2.CAB found at $path." -sColor Green
    }
    catch{
        Write-Log -iTabs 3 "WSUSSCN2.CAB not found at $Path. Script will proceed but won't be able to determine status for updates from MicrosoftUpdate." -sColor Yellow
        $wsusCabStatus = $false
    }
    #endregion
    #region 1.3 Checking if LogType set is functional       
    Write-Log -iTabs 2 "Checking if output settings are valid."
    if (($LogType -eq "RemoteLog") -or ($LogType -eq "Both")){
        Write-Log -iTabs 3 "Testing Remote Log location."
        if (Test-Path $RemoteLogPath){
            Write-Log -iTabs 4 "Remote Log ($RemoteLogPath) location was found." -sColor Green           
        }
        else{
            Write-Log -iTabs 4 "Remote Log location was not found. Remote Logging not possible" -sColor Yellow
        }
    }
    if (($LogType -eq "SQLLog") -or ($LogType -eq "Both")){
        Write-Log -iTabs 3 "Testing SQL Log Settings."        
        $constr = "Server=$SQLServer;Database=$SQLDB;Trusted_Connection=True;"        
        $newConnection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $newConnection.ConnectionString = $constr        
        Write-Log -iTabs 4 "Attempting to connect to the database $constr"           
        try {            
            $newConnection.Close()
            $newConnection.Open()
            Write-Log -iTabs 4 "Connection to database tested successfully." -sColor Green
            $newConnection.Close()
        }
        catch {
            Write-Log -iTabs 4 "Could not open the connection. SQL logging will not be possible" -sColor Yellow                                  
        }        
    }
    #endregion
    Write-Log -iTabs 1 "Completed 1 - Pre-Checks."-sColor Cyan    
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================

# ===============================================================================================================================================================================
#region 2_EXECUTION
    Write-Log -iTabs 1 "Starting 2 - Execution." -sColor cyan    
    $OutPut = @()
    #region 2.1 Get-MissingUpdates from SCCM
    Write-Log -iTabs 2 "Getting missing updates from SCCM."
    If ($sccmUpdateStatus -ne $false){
        Write-Log -iTabs 3 "Querying CCM_UpdatesStore WMI."
        try{
            $sccmMissing = Get-WmiObject -Namespace ROOT\ccm\SoftwareUpdates\UpdatesStore -Query "Select * from CCM_UpdateStatus WHERE Status = 'Missing'" 
            Write-Log -iTabs 4 "$($sccmMissing.Count) updates found missing from WMI." -sColor Green
        }
        catch{
            Write-Log -iTabs 4 "Error retrieving updates from WMI. This is not expected. Aborting script." -sColor Red
            $global:iExitCode = 9001
            return $global:iExitCode
        }
        Write-Log -iTabs 3 "Getting Software Update Group Deployments assigned to this device, from WMI."                
        try{
            $SUGDeployments = Get-WmiObject -Namespace "ROOT\ccm\Policy\Machine\RequestedConfig" -query "Select * FROM CCM_UpdateCIAssignment" | ConvertTo-Array
            Write-Log -iTabs 4 "$($SUGDeployments.Count) deployments found from WMI." -sColor Green
        }
        catch{
            Write-Log -iTabs 4 "Error retrieving deployments from WMI. This is not expected. Aborting script." -sColor Red
            $global:iExitCode = 9002
            return $global:iExitCode
        }    
        Write-Log -iTabs 3 "For each Software Update Group Deployments found, detect Updates Assigned to it."                
        $deployedUpdates=@()
        $count=1
        foreach ($SUGDeployment in $SUGDeployments){
            Write-Log -iTabs 4 "($count/$($SUGDeployments.Count))$($SUGDeployment.AssignmentName) found! Getting Assigned CIs..."                
            foreach ($update in $SUGDeployment.AssignedCIs){
                $UpdateObj=[pscustomobject]@{"DateTime"="";"ComputerName"="";"Source"="";"Article"="";"UpdateID"="";"Title"="";"UpdateClassification"="";"ProductId"="";"Categories"=""}
                [xml]$ofxUpdate = $Update
                $UpdateObj.DateTime = $(Get-Date -UFormat %Y%m%d_%H%M%S)
                $UpdateObj.ComputerName = $env:computername
                $UpdateObj.Source = "SCCM: "+$SUGDeployment.AssignmentName
                try{
                    $UpdateObj.Article = $($ofxUpdate.ci.DisplayName | Select-String -Pattern 'KB\d*' -AllMatches | % { $_.Matches } | % {$_.value}).Replace("KB","")
                    if ($null -eq $UpdateObj.Article){                     
                        $UpdateObj.Article = $($ofxUpdate.ci.DisplayName | Select-String -Pattern ' \(\d*' -AllMatches | % { $_.Matches } | % {$_.value}).Replace(" (","")
                    }
                }
                catch{
                    $UpdateObj.Article = "N/A"
                }                
                $UpdateObj.UpdateID = $ofxUpdate.ci.id           
                $UpdateObj.Title = $ofxUpdate.ci.DisplayName
                $UpdateObj.UpdateClassification = $ofxUpdate.ci.UpdateClassification
                $UpdateObj.ProductId = $ofxUpdate.ci.ApplicabilityCondition.ApplicabilityRule.ProductId
                $deployedUpdates+=$UpdateObj                
            }
            $count++
        }
        Write-Log -iTabs 4 "$($deployedUpdates.count) deployed updates found." -sColor Green
        Write-Log -iTabs 3 "Checking if Missing Updates are found within Deployed Updates." 
        $count=1
        #$deployedUpdates.UpdateID
        foreach ($missingUpdate in $sccmMissing){    
            if ($deployedUpdates.UpdateID -contains $missingUpdate.UniqueID){
                Write-Log -iTabs 4 "Found update missing $($missingUpdate.Article) - $($missingUpdate.Title)" -sColor Red
                $output+=$deployedUpdates | Where {$_.UpdateID -eq $missingUpdate.UniqueID}
            }
            else{
                #Write-Host "$($missingUpdate.UniqueID) not found in deployed updates" -ForegroundColor Green
                
            }                  
            $count++
        }
        Write-Log -iTabs 4 "$($output.count) updates missing from SCCM deployments." -sColor Green
    }
    else{
        Write-Log -iTabs 3 "Query to SCCM updates will not be performed due to pre-check 1.1 failure."
    }
    #endregion
    #region 2.2 Get-MissingUpdates from WSUSSCN2.CAB     
    Write-Log -iTabs 2 "Creating Update Searcher using wsusscn2.cab. This might take a few minutes."
    if ($wsusCabStatus){
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session 
        $UpdateServiceManager  = New-Object -ComObject Microsoft.Update.ServiceManager 
        $UpdateService = $UpdateServiceManager.AddScanPackageService("Offline Sync Service",$Path, 1) 
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()   
        $UpdateSearcher.ServerSelection = 3 #ssOthers 
        $UpdateSearcher.ServiceID = [string]$UpdateService.ServiceID  
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0") # or "IsInstalled=0 and IsInstalled=1" to also list the installed updates as MBSA did  
        Write-Log -iTabs 3 "Retrieving updates from Searcher with Status Missing"
        $Updates = $SearchResult.Updates 
        $OfflineUpdateArray =@()
        Write-Log -iTabs 3 "Building array with updates missing"
        $count = 1
        foreach ($Update in $Updates){       
            $UpdateObj=[pscustomobject]@{"DateTime"="";"ComputerName"="";"Source"="";"Article"="";"UpdateID"="";"Title"="";"UpdateClassification"="";"ProductId"="";"Categories"=""}       
            $UpdateObj.DateTime = $(Get-Date -UFormat %Y%m%d_%H%M%S)
            $UpdateObj.ComputerName = $env:COMPUTERNAME
            $UpdateObj.Source = "WSUSCAB$cabVer"            
            try{
                $UpdateObj.Article = $($Update.Title | Select-String -Pattern 'KB\d*' -AllMatches | % { $_.Matches } | % {$_.value}).Replace("KB","")
                if ($null -eq $UpdateObj.Article){                     
                    $UpdateObj.Article = $($Update.Title | Select-String -Pattern ' \(\d*' -AllMatches | % { $_.Matches } | % {$_.value}).Replace(" (","")
                }
            }
            catch{
                $UpdateObj.Article = "N/A"
            } 
            $UpdateObj.UpdateID = $Update.Identity.UpdateID
            $UpdateObj.Title = $Update.Title
            $catstr=$null
            foreach($cat in $Update.Categories){
                $catstr+=$($cat.Type)+":"+$($cat.Name)+";"
            }             
            $UpdateObj.Categories = $catstr
            $output += $UpdateObj
            $count++
        }
        Write-Log -iTabs 4 "$count updates missing from WSUS CAB." -sColor Green
            
    }
    else{
        Write-Log -iTabs 3 "Query to WSUS updates will not be performed due to pre-check 1.2 failure."
    }
    #endregion    
    #region 2.3 Write Output Locally
        $OutPut | Export-Csv -Path (Join-Path -Path $RemoteLogPath -ChildPath $RemoteLogName) -NoTypeInformation -Append -NoClobber
    #endregion
    #write Output remotely, if chosen
    #write output to sql, if chosen
    Write-Log -iTabs 1 "Completed 2 - Execution." -sColor cyan
    Write-Log -iTabs 0 -bConsole $true
#endregion
# ===============================================================================================================================================================================
        
# ===============================================================================================================================================================================
#region 3_POST-CHECKS
# ===============================================================================================================================================================================
    Write-Log -iTabs 1 "Starting 3 - Post-Checks."-sColor cyan
    #Check if local log was created
    #check if remote log was updates
    #check if SQL was updated
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