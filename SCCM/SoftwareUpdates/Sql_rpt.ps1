
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
    $sContent = "||"+$(Get-Date -UFormat %Y-%m-%d_%H:%M:%S)+"|"+$sTabs  + "|"+$sMessage

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
function Invoke-SQL {
    param(
        [string] $dataSource = ".\SQLEXPRESS",
        [string] $database = "MasterData",
        [string] $sqlCommand = $(throw "Please specify a query.")
      )

    $connectionString = "Data Source=$dataSource; " +
            "Integrated Security=SSPI; " +
            "Initial Catalog=$database"

    $connection = new-object system.data.SqlClient.SQLConnection($connectionString)
    $command = new-object system.data.sqlclient.sqlcommand($sqlCommand,$connection)
    $connection.Open()

    $adapter = New-Object System.Data.sqlclient.sqlDataAdapter $command
    $dataset = New-Object System.Data.DataSet
    $adapter.Fill($dataSet) | Out-Null

    $connection.Close()
    $dataSet.Tables

}
#endregion
# --------------------------------------------------------------------------------------------
#Variables

#region VARIABLES
# Standard Variables
    # *****  Change Logging Path and File Name Here  *****    
    $sOutFileName	= "RptDump.log" # Log File Name    
    $sEventSource   = "RptDump" # Event Source Name
    # ****************************************************
    $sScriptName 	= $MyInvocation.MyCommand
    $sScriptPath 	= Split-Path -Parent $MyInvocation.MyCommand.Path    
    $sLogRoot       = "\\sql01.vlab.varandas.com\RptDump"
    #Check/Create Folder
    $date = get-date -UFormat %Y%m%d
    if (!(Test-Path $($sLogRoot+"\"+$date))){
        try{ New-Item -Path $sLogRoot -ItemType Directory -Name $date } catch{exit}
    }    
    $SLogRoot       = Join-Path -Path $SLogRoot -ChildPath $date    
    $sOutFilePath   = $SLogRoot
    $sLogFile   = $SLogRoot+"\"+$sOutFileName    
  $global:iExitCode = 0
    $sUserName		= $env:username
    $sUserDomain	= $env:userdomain
    $sMachineName	= $env:computername
    $sCMDArgs		= $MyInvocation.Line
    $bAllow64bitRelaunch = $true
    $iLogFileSize 	= 1048576    
    $outputpath = 

Function MainSub{

Write-Log "Starting to get info from SCCM DB" -iTabs 1
#SCCM Query
$dSource = "sccm01.vlab.varandas.com"
Write-Log "SQL Server: $dSource" -iTabs 2
$db = "CM_VAR"
Write-Log "Database: $db" -iTabs 2

    #AllSystems
    $query = "Select 
	fcm.ResourceID 'ResourceID',
	gss.Name0 'GSS_Name',
	gss.SystemRole0 'GSS_Role',
	rsys.Name0 'RSYS_Name',
	rsys.Distinguished_Name0 'RSYS_DN',
	rsys.Client0 'RSYS_ClientInstalled',
	rsys.Client_Version0 'RSYS_ClientVer',
	rsys.Full_Domain_Name0 'RSYS_FullDomain',
	rsys.Operating_System_Name_and0 'RSYS_OSVer',
	gsos.Caption0 'GSOS_OSCaption',
	uss.LastScanTime 'WSUSLastScan',
	uss.LastWUAVersion 'WSUS_WUA_Ver',
	chcs.LastActivetime 'CHCS_LastSCCMActive',
	chcs.ClientStateDescription 'CHCS_SCCMCLientState',
	chcs.LastMPServerName 'CHCS_MPServer'
from 
	v_Collection col
		inner join 
	v_FullCollectionMembership fcm
		on col.CollectionID=fcm.CollectionID
		left join
	v_GS_SYSTEM gss
		on fcm.ResourceID=gss.ResourceID
		left join
	v_R_System rsys
		on fcm.ResourceID=rsys.ResourceID
		left join
	v_GS_OPERATING_SYSTEM gsos
		on fcm.ResourceID=gsos.ResourceID
		left join
	v_UpdateScanStatus uss
		on fcm.ResourceID=uss.ResourceID
		left join
	v_CH_ClientSummary chcs
		on fcm.ResourceID=chcs.ResourceID
		
where col.Name = 'All Systems'
order by 'RSYS_OSVer'"
    Write-Log "Querying All Systems from SCCM" -iTabs 3
    $SCCM_AllSystems = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #AllWorkstations
    $query = "Select 
	fcm.ResourceID 'ResourceID',
	gss.Name0 'GSS_Name',
	gss.SystemRole0 'GSS_Role',
	rsys.Name0 'RSYS_Name',
	rsys.Distinguished_Name0 'RSYS_DN',
	rsys.Client0 'RSYS_ClientInstalled',
	rsys.Client_Version0 'RSYS_ClientVer',
	rsys.Full_Domain_Name0 'RSYS_FullDomain',
	rsys.Operating_System_Name_and0 'RSYS_OSVer',
	gsos.Caption0 'GSOS_OSCaption',
	uss.LastScanTime 'WSUSLastScan',
	uss.LastWUAVersion 'WSUS_WUA_Ver',
	chcs.LastActivetime 'CHCS_LastSCCMActive',
	chcs.ClientStateDescription 'CHCS_SCCMCLientState',
	chcs.LastMPServerName 'CHCS_MPServer'
from 
	v_Collection col
		inner join 
	v_FullCollectionMembership fcm
		on col.CollectionID=fcm.CollectionID
		left join
	v_GS_SYSTEM gss
		on fcm.ResourceID=gss.ResourceID
		left join
	v_R_System rsys
		on fcm.ResourceID=rsys.ResourceID
		left join
	v_GS_OPERATING_SYSTEM gsos
		on fcm.ResourceID=gsos.ResourceID
		left join
	v_UpdateScanStatus uss
		on fcm.ResourceID=uss.ResourceID
		left join
	v_CH_ClientSummary chcs
		on fcm.ResourceID=chcs.ResourceID
		
where col.Name = 'WKS-SUP-DG4-COL'
order by 'RSYS_OSVer'"
    Write-Log "Querying All Workstations from SCCM" -iTabs 3
    $SCCM_AllWorkstations = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #AllUpdatesCompliance
    $query = "
    Select
	fcm.Name,
	ui.Title,
	ui.ArticleID,
	CASE ucsa.Status
		WHEN 0 THEN 'Status Unknown'
		WHEN 1 THEN 'Update Not required'
		WHEN 2 THEN 'Update Missing'
		WHEN 3 THEN 'Update Installed'
	END 'InstallStatus',
	ucsa.LastErrorCode 'LastWUAInstallError',
	CASE ucsa.LastEnforcementMessageID
		WHEN 0 THEN 'Enforcement state unknown' 
		WHEN 1 THEN 'Enforcement started'
		WHEN 2 THEN 'Enforcement waiting for content'
		WHEN 3 THEN 'Waiting for another installation to complete'
		WHEN 4 THEN 'Waiting for maintenance window before installing'
		WHEN 5 THEN 'Restart required before installing'
		WHEN 6 THEN 'General failure'
		WHEN 7 THEN 'Pending installation'
		WHEN 8 THEN 'Installing update'
		WHEN 9 THEN 'Pending system restart'
		WHEN 10 THEN 'Successfully installed update'
		WHEN 11 THEN 'Failed to install update'
		WHEN 12 THEN 'Downloading update'
		WHEN 13 THEN 'Downloaded update'
		WHEN 14 THEN 'Failed to download update'
		END 'LastStatus',
	ucsa.LastEnforcementStatusMsgID 'LastDeploymentError',
	ali.title 'SUG_Title'
from 
	v_Collection col
		inner join 
	v_FullCollectionMembership fcm
		on col.CollectionID=fcm.CollectionID
		left join
	v_Update_ComplianceStatusAll ucsa
		on fcm.ResourceID=ucsa.ResourceID
		left join
	v_UpdateInfo ui
		on ucsa.CI_ID=ui.CI_ID
		left join
	v_CIRelation cir
		on ucsa.CI_ID=cir.ToCIID
		left join
	v_AuthListInfo ali
		on cir.FromCIID=ali.CI_ID
where 
	col.Name = 'WKS-SUP-DG4-COL'
	and ali.IsDeployed = 1
	and ali.Title like 'WKS%'
    "
    Write-Log "Querying All Updates and their compliance from SCCM" -iTabs 3
    $SCCM_AllUpdCompliance= Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #MissingUpdatesCompliance
    $query = "
    Select
	fcm.Name,
	ui.Title,
	ui.ArticleID,
	CASE ucsa.Status
		WHEN 0 THEN 'Status Unknown'
		WHEN 1 THEN 'Update Not required'
		WHEN 2 THEN 'Update Missing'
		WHEN 3 THEN 'Update Installed'
	END 'InstallStatus',
	ucsa.LastErrorCode 'LastWUAInstallError',
	CASE ucsa.LastEnforcementMessageID
		WHEN 0 THEN 'Enforcement state unknown' 
		WHEN 1 THEN 'Enforcement started'
		WHEN 2 THEN 'Enforcement waiting for content'
		WHEN 3 THEN 'Waiting for another installation to complete'
		WHEN 4 THEN 'Waiting for maintenance window before installing'
		WHEN 5 THEN 'Restart required before installing'
		WHEN 6 THEN 'General failure'
		WHEN 7 THEN 'Pending installation'
		WHEN 8 THEN 'Installing update'
		WHEN 9 THEN 'Pending system restart'
		WHEN 10 THEN 'Successfully installed update'
		WHEN 11 THEN 'Failed to install update'
		WHEN 12 THEN 'Downloading update'
		WHEN 13 THEN 'Downloaded update'
		WHEN 14 THEN 'Failed to download update'
		END 'LastStatus',
	ucsa.LastEnforcementStatusMsgID 'LastDeploymentError',
	ali.title 'SUG_Title'
from 
	v_Collection col
		inner join 
	v_FullCollectionMembership fcm
		on col.CollectionID=fcm.CollectionID
		left join
	v_Update_ComplianceStatusAll ucsa
		on fcm.ResourceID=ucsa.ResourceID
		left join
	v_UpdateInfo ui
		on ucsa.CI_ID=ui.CI_ID
		left join
	v_CIRelation cir
		on ucsa.CI_ID=cir.ToCIID
		left join
	v_AuthListInfo ali
		on cir.FromCIID=ali.CI_ID
where 
	col.Name = 'All Workstations'
	and ali.IsDeployed = 1
	and ali.Title like 'WKS%'
	and ucsa.Status in (0,2)

    "
    Write-Log "Querying Missing Updates and their compliance from SCCM" -iTabs 3
    $SCCM_MissingUpdCompliance= Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

#WSUS Query
Write-Log "Starting to get info from WSUS DB" -iTabs 1
$dSource = "sccm01.vlab.varandas.com"
Write-Log "SQL Server: $dSource" -iTabs 2
$db = "SUSDB"
Write-Log "Database: $db" -iTabs 2

    #AllSystems
    $query = "SELECT [ComputerTargetId]
          ,[Name]
          ,[IPAddress]
          ,[LastSyncTime]
          ,[ClientVersion]
          ,[Make]
          ,[Model]      
          ,[OSDefaultUILanguage]
      FROM [SUSDB].[PUBLIC_VIEWS].[vComputerTarget]"
    Write-Log "Querying All Systems from WSUS" -iTabs 3
    $WSUS_AllSystems = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #AllCategories
    $query = "select DefaultTitle Category
    from [SUSDB].[PUBLIC_VIEWS].[vCategory]
    where
    categorytype ='Product'
    order by 1" 
    Write-Log "Querying All Categories from WSUS" -iTabs 3
    $WSUS_AllCategory = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #AllTrackedUpdates
    $query = "
    SELECT
	     upd.[KnowledgebaseArticle]  'KB_Article'
	    ,upd.[DefaultTitle] 'KB_Title'
	    ,cla.DefaultTitle 'KB_Classification'  
	    ,cat.DefaultTitle   'KB_Category' 
	    ,[CreationDate]      
    FROM 
	              [SUSDB].[PUBLIC_VIEWS].[vUpdate] upd
	    left join [SUSDB].[PUBLIC_VIEWS].[vClassification] cla
		    on upd.ClassificationID=cla.ClassificationID
	    left join [SUSDB].[PUBLIC_VIEWS].[vUpdateInCategory] UiC
		    on upd.UpdateId=uic.UpdateID
	    left join [SUSDB].[PUBLIC_VIEWS].[vCategory] cat
		    on uic.CategoryId=cat.CategoryId
    Where 
	        cla.DefaultTitle = 'Security Updates'
	    and upd.IsDeclined=0
	    and	uic.CategoryType = 'Product'
	    and	cat.DefaultTitle in (		
		    'Office 2003',
		    'Office 2007',
		    'Office 2010',
		    'Office 2013',		
		    'Windows 10',		
		    'Windows 10 LTSB',		
		    'Windows 7',		
		    'Windows 8.1'		
	    )
	    and (
			    upd.DefaultTitle not like '%Beta%'
		    and upd.DefaultTitle not like '%Windows 8.1 for x86-based Systems (KB%'
		    and upd.DefaultTitle not like '%Windows 8.1 (KB%'
		    and upd.DefaultTitle not like '%Windows 10 Version Next%'
		    and upd.DefaultTitle not like '%ARM64-based%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1507%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1703%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1709%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1803%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1903%'
		    and upd.DefaultTitle not like '%Quality Rollup%'
		    and upd.DefaultTitle not like '%Productivity%'
		    and upd.DefaultTitle not like '%Learning Essentials%'
		    and upd.DefaultTitle not like '%Chart Controls%'
		    and upd.DefaultTitle not like '%Web App%'
		    and upd.DefaultTitle not like '%FAST Search%'
		    and upd.DefaultTitle not like '%Filter Pack%'
		    and upd.DefaultTitle not like '%Groove%'
		    and upd.DefaultTitle not like '%Office Forms Server%'
		    and upd.DefaultTitle not like '%Sharepoint%'
		    and upd.DefaultTitle not like '%Project Server%'
		    and upd.DefaultTitle not like '%Small Business%'
		    and upd.DefaultTitle not like '%Multilingual%'
		    and upd.DefaultTitle not like '%Proofing%'
		    and upd.DefaultTitle not like '%Windows 8.1 Update (KB%'
	    )
    "
    Write-Log "Querying All Tracked Updates from WSUS" -iTabs 3
    $WSUS_AllTrackedUpdates = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #AllUpdateCompliance
    $query = "
    SELECT 
	    comp.[Name] Device,
	    upd.[DefaultTitle] KBTitle,      
	    upd.[KnowledgebaseArticle] KBArticle,
	    case uii.State
		    WHEN 0 THEN 'Unknown' 
		    WHEN 1 THEN 'NotApplicable' 
		    When 2 THEN 'NotInstalled' 
		    WHEN 3 THEN 'Downloaded' 
		    WHEN 4 THEN 'Installed' 
		    WHEN 5 THEN 'Failed' 
		    WHEN 6 THEN 'InstalledPendingReboot' 
		    Else 'No Match'
	    END   KBState	  
    FROM 
			      [SUSDB].[PUBLIC_VIEWS].[vUpdate] upd
	    left join [SUSDB].[PUBLIC_VIEWS].[vClassification] cla
		    on upd.ClassificationID=cla.ClassificationID
	    left join [SUSDB].[PUBLIC_VIEWS].[vUpdateInCategory] UiC
		    on upd.UpdateId=uic.UpdateID
	    left join [SUSDB].[PUBLIC_VIEWS].[vCategory] cat
		    on uic.CategoryId=cat.CategoryId
	    left join [SUSDB].[PUBLIC_VIEWS].[vUpdateInstallationInfo] uii
		    on upd.UpdateId=uii.UpdateId
	    left join [SUSDB].[PUBLIC_VIEWS].[vComputerTarget] comp
		    on uii.ComputerTargetId=comp.ComputerTargetId
    Where 
		    cla.DefaultTitle = 'Security Updates'
	    and upd.IsDeclined=0
	    and uic.CategoryType = 'Product'
	    and cat.DefaultTitle in (
		    'Office 2003',
		    'Office 2007',
		    'Office 2010',
		    'Office 2013',		
		    'Windows 10',		
		    'Windows 10 LTSB',		
		    'Windows 7',		
		    'Windows 8.1'		
	    )
	    and 
	    (		upd.DefaultTitle not like '%Beta%'
		    and upd.DefaultTitle not like '%Windows 8.1 for x86-based Systems (KB%'
		    and upd.DefaultTitle not like '%Windows 8.1 (KB%'
		    and upd.DefaultTitle not like '%Windows 10 Version Next%'
		    and upd.DefaultTitle not like '%ARM64-based%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1507%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1703%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1709%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1803%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1903%'
		    and upd.DefaultTitle not like '%Quality Rollup%'
		    and upd.DefaultTitle not like '%Productivity%'
		    and upd.DefaultTitle not like '%Learning Essentials%'
		    and upd.DefaultTitle not like '%Chart Controls%'
		    and upd.DefaultTitle not like '%Web App%'
		    and upd.DefaultTitle not like '%FAST Search%'
		    and upd.DefaultTitle not like '%Filter Pack%'
		    and upd.DefaultTitle not like '%Groove%'
		    and upd.DefaultTitle not like '%Office Forms Server%'
		    and upd.DefaultTitle not like '%Sharepoint%'
		    and upd.DefaultTitle not like '%Project Server%'
		    and upd.DefaultTitle not like '%Small Business%'
		    and upd.DefaultTitle not like '%Multilingual%'
		    and upd.DefaultTitle not like '%Proofing%'
		    and upd.DefaultTitle not like '%Windows 8.1 Update (KB%'
	    )
    "
    Write-Log "Querying All Updates Compliance from WSUS" -iTabs 3
    $WSUS_UpdateCompliance = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

    #MissingUpdateCompliance
    $query = "
    SELECT 
	    comp.[Name] Device,
	    upd.[DefaultTitle] KBTitle,      
	    upd.[KnowledgebaseArticle] KBArticle,
	    case uii.State
		    WHEN 0 THEN 'Unknown' 
		    WHEN 1 THEN 'NotApplicable' 
		    When 2 THEN 'NotInstalled' 
		    WHEN 3 THEN 'Downloaded' 
		    WHEN 4 THEN 'Installed' 
		    WHEN 5 THEN 'Failed' 
		    WHEN 6 THEN 'InstalledPendingReboot' 
		    Else 'No Match'
	    END   KBState	  
    FROM 
			      [SUSDB].[PUBLIC_VIEWS].[vUpdate] upd
	    left join [SUSDB].[PUBLIC_VIEWS].[vClassification] cla
		    on upd.ClassificationID=cla.ClassificationID
	    left join [SUSDB].[PUBLIC_VIEWS].[vUpdateInCategory] UiC
		    on upd.UpdateId=uic.UpdateID
	    left join [SUSDB].[PUBLIC_VIEWS].[vCategory] cat
		    on uic.CategoryId=cat.CategoryId
	    left join [SUSDB].[PUBLIC_VIEWS].[vUpdateInstallationInfo] uii
		    on upd.UpdateId=uii.UpdateId
	    left join [SUSDB].[PUBLIC_VIEWS].[vComputerTarget] comp
		    on uii.ComputerTargetId=comp.ComputerTargetId
    Where 
		    cla.DefaultTitle = 'Security Updates'
	    and upd.IsDeclined=0
        and uii.State!=4
	    and uic.CategoryType = 'Product'
	    and cat.DefaultTitle in (
		    'Office 2003',
		    'Office 2007',
		    'Office 2010',
		    'Office 2013',		
		    'Windows 10',		
		    'Windows 10 LTSB',		
		    'Windows 7',		
		    'Windows 8.1'		
	    )
	    and 
	    (		upd.DefaultTitle not like '%Beta%'
		    and upd.DefaultTitle not like '%Windows 8.1 for x86-based Systems (KB%'
		    and upd.DefaultTitle not like '%Windows 8.1 (KB%'
		    and upd.DefaultTitle not like '%Windows 10 Version Next%'
		    and upd.DefaultTitle not like '%ARM64-based%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1507%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1703%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1709%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1803%'
		    and upd.DefaultTitle not like '%Windows 10 Version 1903%'
		    and upd.DefaultTitle not like '%Quality Rollup%'
		    and upd.DefaultTitle not like '%Productivity%'
		    and upd.DefaultTitle not like '%Learning Essentials%'
		    and upd.DefaultTitle not like '%Chart Controls%'
		    and upd.DefaultTitle not like '%Web App%'
		    and upd.DefaultTitle not like '%FAST Search%'
		    and upd.DefaultTitle not like '%Filter Pack%'
		    and upd.DefaultTitle not like '%Groove%'
		    and upd.DefaultTitle not like '%Office Forms Server%'
		    and upd.DefaultTitle not like '%Sharepoint%'
		    and upd.DefaultTitle not like '%Project Server%'
		    and upd.DefaultTitle not like '%Small Business%'
		    and upd.DefaultTitle not like '%Multilingual%'
		    and upd.DefaultTitle not like '%Proofing%'
		    and upd.DefaultTitle not like '%Windows 8.1 Update (KB%'
	    )
    "
    Write-Log "Querying Missing Updates Compliance from WSUS" -iTabs 3
    $WSUS_MissingUpdateCompliance = Invoke-SQL -dataSource $dSource -database $db -sqlCommand $query

#Writting Output
Write-Log "Exporting CSVs" -iTabs 1
Write-Log "Exporting SCCM_AllSystems.csv to $SLogRoot" -iTabs 2
$SCCM_AllSystems | Export-CSV $($SLogRoot+"\SCCM_AllSystems.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting SCCM_AllWorkstations.csv to $SLogRoot" -iTabs 2
$SCCM_AllWorkstations | Export-CSV $($SLogRoot+"\SCCM_AllWorkstations.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting SCCM_AllUpdCompliance.csv to $SLogRoot" -iTabs 2
$SCCM_AllUpdCompliance | Export-CSV $($SLogRoot+"\SCCM_AllUpdCompliance.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting SCCM_MissingUpdCompliance.csv" -iTabs 2
$SCCM_MissingUpdCompliance | Export-CSV $($SLogRoot+"\SCCM_MissingUpdCompliance.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting WSUS_AllSystems.csv to $SLogRoot" -iTabs 2
$WSUS_AllSystems | Export-CSV $($SLogRoot+"\WSUS_AllSystems.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting WSUS_AllCategory.csv to $SLogRoot" -iTabs 2
$WSUS_AllCategory | Export-CSV $($SLogRoot+"\WSUS_AllCategory.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting WSUS_AllTrackedUpdates.csv to $SLogRoot" -iTabs 2
$WSUS_AllTrackedUpdates | Export-CSV $($SLogRoot+"\WSUS_AllTrackedUpdates.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting WSUS_UpdateCompliance.csv to $SLogRoot" -iTabs 2
$WSUS_UpdateCompliance | Export-CSV $($SLogRoot+"\WSUS_UpdateCompliance.csv") -Force -Append -NoTypeInformation
Write-Log "Exporting WSUS_MissingUpdateCompliance.csv to $SLogRoot" -iTabs 2
$WSUS_MissingUpdateCompliance | Export-CSV $($SLogRoot+"\WSUS_MissingUpdateCompliance.csv") -Force -Append -NoTypeInformation


}
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