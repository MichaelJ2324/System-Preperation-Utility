#Region Includes
;*****Includes ********
;Includes for Various functions built in to AutoIt.
;Copy.au3 is the reason copying drivers has a progress bar.
#include <GuiConstantsEx.au3>
#include <GuiConstants.au3> 
#include <WindowsConstants.au3>
#include <AVIConstants.au3>
#include <File.au3>
#include <Constants.au3>
#include <Array.au3>
#include <Copy.au3>
#include <ErrorInterpreter.au3>
#include <Date.au3>

#EndRegion Includes
Global $fromSysprep=""
if ($CmdLine[0]>0) Then
   $fromSysprep=$CmdLine[1]
EndIf
Init()
LoadGUI()


#Region Initialize
Func Init()
   Global $sysprepConf="sysprep_config.ini"
   
   ;Load sysprep Type from config
   Global $sysprepType = IniRead($sysprepConf,"Config","sysprepType","Network")
   Global $sysprepUSBDrive = IniRead($sysprepConf,"Config","usbDriveLabel","")
   
   ;Load sysprep config from INI
   Global $dataIniFile = IniRead($sysprepConf,"Config","dataFile","data.ini")
   Global $dataIniPath = IniRead($sysprepConf,"Config","dataPath","sysinfo")
   ;Load sysprep logging config from INI
   Global $g_LogFileDirectory = IniRead($sysprepConf,"Config","logFileDir","LogFiles")
   Global $g_LogFileName = IniRead($sysprepConf,"Config","logFileName","<serviceTag>_sysprepLog.txt")
   Global $g_LogLevel = IniRead($sysprepConf,"Config","logLevel","5")
   
   ;Load sysprep server configuration from INI
   ;;Check if using global server
   Global $useGlobalServer = IniRead($sysprepConf,"Servers","useGlobal","True")

   ;;Load all servers seperately	  
   Global $globalServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
   Global $backupServer = IniRead($sysprepConf,"Servers","backupServer","ittest")
   Global $usmtServer = IniRead($sysprepConf,"Servers","usmtServer","ittest")
   Global $sysprepServer = IniRead($sysprepConf,"Servers","sysprepServer","ittest")
   Global $driversServer = IniRead($sysprepConf,"Servers","driversServer","ittest")
   Global $imageXServer = IniRead($sysprepConf,"Servers","imageXServer","ittest")

   ;;Load Username and Passwords for servers
   Global $globalUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
   Global $globalPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
   Global $backupUser = IniRead($sysprepConf,"Servers","backupUser","ITTest\sysprep")
   Global $backupPassword = IniRead($sysprepConf,"Servers","backupPassword","FWM$sysprep24")
   Global $usmtUser = IniRead($sysprepConf,"Servers","usmtUser","ITTest\sysprep")
   Global $usmtPassword = IniRead($sysprepConf,"Servers","usmtPassword","FWM$sysprep24")
   Global $sysprepUser = IniRead($sysprepConf,"Servers","sysprepUser","ITTest\sysprep")
   Global $sysprepPassword = IniRead($sysprepConf,"Servers","sysprepPassword","FWM$sysprep24")
   Global $driversUser = IniRead($sysprepConf,"Servers","driversUser","ITTest\sysprep")
   Global $driversPassword = IniRead($sysprepConf,"Servers","driversPassword","FWM$sysprep24")
   Global $imageXUser = IniRead($sysprepConf,"Servers","imageXUser","ITTest\sysprep")
   Global $imageXPassword = IniRead($sysprepConf,"Servers","imageXPassword","FWM$sysprep24")
   
   ;Load backup configuration
   ;;Backup path on server, drive letter, and filename
   Global $backupFilePath = IniRead($sysprepConf,"Backup","backupPath","data\Backups")
   Global $backupDrive = IniRead($sysprepConf,"Backup","backupDrive","P")
   Global $backupFilename = IniRead($sysprepConf,"Backup","backupFilename","<userName> <modelNum> <mon>-<day>-<year>")
   GLobal $backupLogFile = IniRead($sysprepConf,"Backup","backupLogFile","<serviceTag>_backup.log")
   ;Load USMT configuration
   ;;USMT path on server and drive letter
   Global $usmtFilePath = IniRead($sysprepConf,"USMT","usmtPath","data\Backups\USMT")
   Global $usmtDrive = IniRead($sysprepConf,"USMT","usmtDrive","U")
   Global $usmtTempDir = IniRead($sysprepConf,"USMT","usmtTempDir","Temp<rand>")
   Global $usmtXMLFile = IniRead($sysprepConf,"USMT","usmtXMLFile","usermigrate.xml")
   GLobal $usmtLogFile = IniRead($sysprepConf,"USMT","usmtLogFile","<serviceTag>_scanstate.log")
   Global $usmtListFiles = IniRead($sysprepConf,"USMT","usmtListFiles","True")
   Global $usmtListFilesFileName = IniRead($sysprepConf,"USMT","listFilesFileName","<serviceTag>_files.txt")
   ;Load Sysprep Configuration
   ;;sysprep path on server and drive leter
   Global $sysprepFilePath = IniRead($sysprepConf,"Config","sysprepPath","data\Sysprep")
   Global $sysprepDrive = IniRead($sysprepConf,"Config","sysprepDrive","W")
   ;Load Driver Server configuration
   Global $driversFilePath = IniRead($sysprepConf,"Drivers","driversPath","data\Drivers")
   Global $driversDrive = IniRead($sysprepConf,"Drivers","driversDrive","Y")
   ;Load ImageX Configuration
   Global $imageXFilePath = IniRead($sysprepConf,"ImageX","imageXPath","data\images\current")
   Global $imageXDrive = IniRead($sysprepConf,"ImageX","imageXDrive","I")
   Global $defaultWIM = IniRead($sysprepConf,"ImageX","defaultWIM","current.wim")
   
   ;Open file for writing
   ;;Using mode 8 because even it it doesn't exist this will create it
   Global $configFile = FileOpen($sysprepConf,8)
   
   Global $savedConfig=False
EndFunc
#endregion Initialize

#Region GUI
Func LoadGUI()
   Opt("GUIOnEventMode", 1)
   $configMgrForm = GuiCreate("FWM SysPrep - Configuration", 700, 500)
	  ;Create Menus in top left corner
	  ;;File Menu Contains: 
	  ;;;	Open Config File - Opens config file in notepad for editing
	  ;;;	Load Data File - Just in case you clicked no at the first prompt, or if you edited the data file you can reload
	  ;;;					the variables back into the program.
	  ;;;	Exit - Exit
	  $fileMenu=GuiCtrlCreateMenu("File")
		 $dataOpenItem = GuiCtrlCreateMenuItem("Open Config File", $fileMenu)
		 $sysPrepLog=GuiCtrlCreateMenuItem("View SysPrep Log", $fileMenu)
		 $exitItem=GuiCtrlCreateMenuItem("Exit", $fileMenu)	
	  $helpMenu=GUICtrlCreateMenu("Help")
		 $instructionsButton = GuiCtrlCreateMenuItem("Instructions",$helpMenu)
	  $saveAll = GUICtrlCreateButton("Save All", 10, 10, 100, 35)
	  
	  GUICtrlCreateTab(0,50,702,452)
	  ;;Creates tab structure and menues on each Tab in order displayed
	  
	  	  ;;Sysprep Tab
	  $globalTab = GUICtrlCreateTabItem("Sysprep")
		 $title1 = GUICtrlCreateLabel("Sysprep Settings",7, 75)
		 
		 $sysprepTypeLabel = GUICtrlCreateLabel("Sysprep Type:",7,93)
		 Global $sysprepTypeCombo = GUICtrlCreateCombo("Network",100,90,125,20)
		 $sysprepTypeData = GUICtrlSetData($sysprepTypeCombo,"USB",$sysprepType)
		 $sysprepUSBDriveLabel = GuiCtrlCreateLabel("USB Drive Label:",7,123)
		 Global $sysprepUSBDriveInput = GUICtrlCreateInput($sysprepUSBDrive,100,120,250,20)
		 
		 $usmtDriveLabel = GUICtrlCreateLabel("Drive Letter:",7,173)
		 Global $sysprepDriveInput = GUICtrlCreateInput($sysprepDrive,100,170,125,20)
		 $usmtPathLabel = GUICtrlCreateLabel("Path On Server:",7,203)
		 Global $sysprepPathInput = GUICtrlCreateInput($sysprepFilePath,100,200,250,20)
		 $dataFilePathLabel = GUICtrlCreateLabel("Data File Path:",7,233)
		 Global $dataFilePathInput = GUICtrlCreateInput($dataIniPath,100,230,250,20)
		 $dataFileLabel = GUICtrlCreateLabel("Data File Name:",7,263)
		 Global $dataFileInput = GUICtrlCreateInput($dataIniFile,100,260,250,20)

		 $logFilePathLabel = GUICtrlCreateLabel("Log File Path:",7,293)
		 Global $logFilePathInput = GUICtrlCreateInput($g_logFileDirectory,100,290,250,20)
		 $logFileNameLabel = GUICtrlCreateLabel("Log File Name:",7,323)
		 Global $logFileNameInput = GUICtrlCreateInput($g_logFileName,100,320,250,20)
		 $logLevelLabel = GUICtrlCreateLabel("Log Level (1-5):",7,353)
		 Global $logLevelInput = GUICtrlCreateInput($g_logLevel,100,350,125,20)
		 
		 Global $saveGlobalButton = GUICtrlCreateButton("Save Global Settings",7,420,150,40)
	  
	  ;;Server Tab
	  $serverTab = GUICtrlCreateTabItem("Servers")
		 $title5 = GuiCtrlCreateLabel("Server Configuration", 10, 75)
		 Global $useGlobalCheckBox = GUICtrlCreateCheckbox("Use Global Server?", 7, 93)

		 $globalServerLabel = GUICtrlCreateLabel("Global Server:",7,123)
		 Global $globalServerInput = GUICtrlCreateInput($globalServer,100,120,125,20)
		 $globalUserLabel = GUICtrlCreateLabel("Username:",7,153)
		 Global $globalUserInput = GUICtrlCreateInput($globalUser,100,150,125,20)
		 $globalPasswordLabel = GUICtrlCreateLabel("Password:",7,183)
		 Global $globalPasswordInput = GUICtrlCreateInput($globalPassword,100,180,125,20,0x0020)
		 
		 $backupServerLabel = GUICtrlCreateLabel("Backup Server:",235,123)
		 Global $backupServerInput = GUICtrlCreateInput($backupServer,327,120,125,20)
		 $backupUserLabel = GUICtrlCreateLabel("Username:",235,153)
		 Global $backupUserInput = GUICtrlCreateInput($backupUser,327,150,125,20)
		 $backupPasswordLabel = GUICtrlCreateLabel("Password:",235,183)
		 Global $backupPasswordInput = GUICtrlCreateInput($backupPassword,327,180,125,20,0x0020)
		 
		 $usmtServerLabel = GUICtrlCreateLabel("USMT Server:",235,223)
		 Global $usmtServerInput = GUICtrlCreateInput($usmtServer,327,220,125,20)
		 $usmtUserLabel = GUICtrlCreateLabel("Username:",235,253)
		 Global $usmtUserInput = GUICtrlCreateInput($usmtUser,327,250,125,20)
		 $usmtPasswordLabel = GUICtrlCreateLabel("Password:",235,283)
		 Global $usmtPasswordInput = GUICtrlCreateInput($usmtPassword,327,280,125,20,0x0020)
		 
		 $sysprepServerLabel = GUICtrlCreateLabel("Sysprep Server:",460,123)
		 Global $sysprepServerInput = GUICtrlCreateInput($sysprepServer,553,120,125,20)
		 $sysprepUserLabel = GUICtrlCreateLabel("Username:",460,153)
		 Global $sysprepUserInput = GUICtrlCreateInput($sysprepUser,553,150,125,20)
		 $sysprepPasswordLabel = GUICtrlCreateLabel("Password:",460,183)
		 Global $sysprepPasswordInput = GUICtrlCreateInput($sysprepPassword,553,180,125,20,0x0020)
		 
		 $driversServerLabel = GUICtrlCreateLabel("Drivers Server:",460,223)
		 Global $driversServerInput = GUICtrlCreateInput($driversServer,553,220,125,20)
		 $driversUserLabel = GUICtrlCreateLabel("Username:",460,253)
		 Global $driversUserInput = GUICtrlCreateInput($driversUser,553,250,125,20)
		 $driversPasswordLabel = GUICtrlCreateLabel("Password:",460,283)
		 Global $driversPasswordInput = GUICtrlCreateInput($driversPassword,553,280,125,20,0x0020)
		 
		 $imageXServerLabel = GUICtrlCreateLabel("Image Server:",235,323)
		 Global $imageXServerInput = GUICtrlCreateInput($imageXServer,327,320,125,20)
		 $imageXUserLabel = GUICtrlCreateLabel("Username:",235,353)
		 Global $imageXUserInput = GUICtrlCreateInput($imageXUser,327,350,125,20)
		 $imageXPasswordLabel = GUICtrlCreateLabel("Password:",235,383)
		 Global $imageXPasswordInput = GUICtrlCreateInput($imageXPassword,327,380,125,20,0x0020)
		 
		 Global $saveServerButton = GUICtrlCreateButton("Save Server Settings",7,420,150,40)
		 
		 If $useGlobalServer=True Then
			GUICtrlSetState($useGlobalCheckBox,$GUI_CHECKED)
			_Server_Input_Toggle()
		 EndIf
	  
	  ;;Backup tab
	  $backupTab = GUICtrlCreateTabItem("Backup")
		 $title3 = GUICtrlCreateLabel("Backup Configuration", 10, 75)
		 
		 $backupDriveLabel = GUICtrlCreateLabel("Drive Letter:",7,93)
		 Global $backupDriveInput = GUICtrlCreateInput($backupDrive,100,90,125,20)
		 $backupPathLabel = GUICtrlCreateLabel("Path On Server:",7,123)
		 Global $backupPathInput = GUICtrlCreateInput($backupFilePath,100,120,250,20)
		 $backupFileNameLabel = GUICtrlCreateLabel("Backup File Name:",7,153)
		 Global $backupFileNameInput = GUICtrlCreateInput($backupFilename,100,150,250,20)
		 $backupLogFileLabel = GUICtrlCreateLabel("Backup Log File:",7,183)
		 Global $backupLogFileInput = GUICtrlCreateInput($backupLogFile,100,180,250,20)
		 
		 Global $saveBackupButton = GUICtrlCreateButton("Save Backup Settings",7,420,150,40)
		 
	  ;;USMT Tab
	  $usmtTab = GUICtrlCreateTabItem("USMT")
		 $title4 = GUICtrlCreateLabel("USMT Configuration", 10, 75)
		 
		 $usmtDriveLabel = GUICtrlCreateLabel("Drive Letter:",7,93)
		 Global $usmtDriveInput = GUICtrlCreateInput($usmtDrive,100,90,125,20)
		 $usmtPathLabel = GUICtrlCreateLabel("Path on Server:",7,123)
		 Global $usmtPathInput = GUICtrlCreateInput($usmtFilePath,100,120,250,20)
		 $usmtTempDirLabel = GUICtrlCreateLabel("Temp Directory:",7,153)
		 Global $usmtTempDirInput = GUICtrlCreateInput($usmtTempDir,100,150,250,20)
		 $usmtXMLFileLabel = GUICtrlCreateLabel("USMT XML File:",7,183)
		 Global $usmtXMLFileInput = GUICtrlCreateInput($usmtXMLFile,100,180,250,20)
		 $usmtLogFileLabel = GUICtrlCreateLabel("USMT Log File:",7,213)
		 Global $usmtLogFileInput = GUICtrlCreateInput($usmtLogFile,100,210,250,20)
		 Global $usmtListFilesInput = GUICtrlCreateCheckbox("List Files Log:",7,243)
		 $usmtListFileFilenameLabel = GUICtrlCreateLabel("Filename: ",7,273)
		 Global $usmtListFilesFileNameInput = GUICtrlCreateInput($usmtListFilesFileName,100,270,250,20)
		 Global $saveUsmtButton = GUICtrlCreateButton("Save USMT Settings",7,420,150,40)
		 
		 If $usmtListFiles=True Then
			GUICtrlSetState($usmtListFilesInput,$GUI_CHECKED)
			_USMT_ListFiles_Input_Toggle()
		 EndIf
	  ;;Drivers Tab
	  $driversTab = GUICtrlCreateTabItem("Drivers")
		 $title8 = GUICtrlCreateLabel("Driver Configuration", 10, 75)
		 
		 $driversDriveLabel = GUICtrlCreateLabel("Drive Letter:",7,93)
		 Global $driversDriveInput = GUICtrlCreateInput($driversDrive,100,90,125,20)
		 $driversPathLabel = GUICtrlCreateLabel("Path On Server:",7,123)
		 Global $driversPathInput = GUICtrlCreateInput($driversFilePath,100,120,250,20)
		 
		 Global $saveDriversButton = GUICtrlCreateButton("Save Driver Settings",7,420,150,40)
	  
	  ;;ImageX
	  $imageXTab = GuiCtrlCreateTabItem("ImageX")
		 $title6 = GuiCtrlCreateLabel("ImageX86 Configuration", 10, 75)
		 
		 $imageXDriveLabel = GUICtrlCreateLabel("Drive Letter:",7,93)
		 Global $imageXDriveInput = GUICtrlCreateInput($imageXDrive,100,90,125,20)
		 $imageXPathLabel = GUICtrlCreateLabel("Path On Server:",7,123)
		 Global $imageXPathInput = GUICtrlCreateInput($imageXFilePath,100,120,250,20)
		 $defaultWIMLabel = GUICtrlCreateLabel("Default WIM:",7,153)
		 Global $defaultWIMInput = GUICtrlCreateInput($defaultWIM,100,150,250,20)
		 
		 Global $saveImageXButton = GUICtrlCreateButton("Save ImageX Settings",7,420,150,40)
   
	  ;$advancedTab = GUICtrlCreateTabItem("Advanced")
		 ;$title10 = GuiCtrlCreateLabel("Advanced Configration", 10, 75)
   
   
   _Sysprep_Type_Toggle()
   ;Add event handlers
   ;;Event handler for Save all button 
   GUICtrlSetOnEvent($saveAll, "_Save_All")
   GUICtrlSetOnEvent($saveGlobalButton,"_Save_Global")
   GUICtrlSetOnEvent($saveServerButton,"_Save_Servers")
   GUICtrlSetOnEvent($saveBackupButton,"_Save_Backup")
   GUICtrlSetOnEvent($saveUsmtButton,"_Save_USMT")
   GUICtrlSetOnEvent($saveDriversButton,"_Save_Drivers")
   GUICtrlSetOnEvent($saveImageXButton,"_Save_ImageX")
   
   ;Event handler for global server disabling/enabling of inputs
   GUICtrlSetOnEvent($useGlobalCheckBox,"_Server_Input_Toggle")
   ;Event handler for List Files toggle
   GUICtrlSetOnEvent($usmtListFilesInput,"_USMT_ListFiles_Input_Toggle")
   ;Sysprep Type event handler
   GUICtrlSetOnEvent($sysprepTypeCombo,"_Sysprep_Type_Toggle")
   
   GUICtrlSetOnEvent($dataOpenItem, "_Open_Config")
   GUICtrlSetOnEvent($instructionsButton,"_Instructions")
   GUICtrlSetOnEvent($exitItem, "CloseWindow")
   GUISetOnEvent($GUI_EVENT_CLOSE, "CloseWindow")
      
   ;Set GUI to show
   GUISwitch($configMgrForm)
   GUISetState(@SW_SHOW,$configMgrForm)																										
   While 1
	   Sleep(1000)
   WEnd
EndFunc
#endregion

#Region Write Configuration
Func _Save_All()
   ProgressOn("Saving Configuration","Saving Configuration", "Please wait...")
   Sleep(600)
   $progress = _Save_Global()
   ProgressSet($progress)
   Sleep(900)
   
   $progress += _Save_Servers()
   ProgressSet($progress)
   Sleep(750)
   $progress += _Save_Backup()
   ProgressSet($progress)
   Sleep(600)
   $progress += _Save_USMT()
   ProgressSet($progress)
   Sleep(300)
   $progress += _Save_Drivers()
   ProgressSet($progress)
   Sleep(300)
   $progress += _Save_ImageX()
   ProgressSet($progress, "Complete")
   Sleep(100)
   ProgressOff()
   $savedConfig=True
EndFunc
Func _Save_Global()
   $sysprepType = GuiCtrlread($sysprepTypeCombo)
   $sysprepUSBDrive = GuiCtrlread($sysprepUSBDriveInput)
   $dataIniFile=GuiCtrlRead($dataFileInput)
   $dataIniPath=GuiCtrlRead($dataFilePathInput)
   $sysprepDrive=GuiCtrlRead($sysprepDriveInput)
   $sysprepFilePath=GuiCtrlRead($sysprepPathInput)
   $g_LogFileDirectory=GuiCtrlRead($logFilePathInput)
   $g_LogFileName=GuiCtrlRead($logFileNameInput)
   $g_LogLevel=GuiCtrlRead($logLevelInput)
   
   IniWrite($sysprepConf,"Config","sysprepType",$sysprepType)
   IniWrite($sysprepConf,"Config","usbDriveLabel",$sysprepUSBDrive)
   IniWrite($sysprepConf,"Config","sysprepPath",$sysprepFilePath)
   If ($sysprepType='USB') Then
	  $sysprepDrive = ""
   EndIf
   IniWrite($sysprepConf,"Config","sysprepDrive",$sysprepDrive)
   IniWrite($sysprepConf,"Config","sysprepPath",$sysprepFilePath)
   IniWrite($sysprepConf,"Config","dataPath",$dataIniPath)
   IniWrite($sysprepConf,"Config","dataFile",$dataIniFile)
   IniWrite($sysprepConf,"Config","logFileDir",$g_LogFileDirectory)
   IniWrite($sysprepConf,"Config","logFileName",$g_LogFileName)
   IniWrite($sysprepConf,"Config","logLevel",$g_LogLevel)
   
   Return 20
EndFunc
Func _Save_Servers()
   If GUICtrlRead($useGlobalCheckBox) = $GUI_CHECKED Then
	  $useGlobalServer=True
	  $globalServer=GuiCtrlRead($globalServerInput)
	  $globalUser=GuiCtrlRead($globalUserInput)
	  $globalPassword=GuiCtrlRead($globalPasswordInput)
	  
	  $backupServer=""
	  $backupUser=""
	  $backupPassword=""
	  
	  $usmtServer=""
	  $usmtUser=""
	  $usmtPassword=""
	  
	  $sysprepServer=""
	  $sysprepUser=""
	  $sysprepPassword=""
	  
	  $driversServer=""
	  $driversUser=""
	  $driversPassword=""
	  
	  $imageXServer=""
	  $imageXUser=""
	  $imageXPassword=""
	  
   Else
	  $useGlobalServer=False
	  $globalServer=""
	  $globalUser=""
	  $globalPassword=""
	  
	  $backupServer=GuiCtrlRead($backupServerInput)
	  $backupUser=GuiCtrlRead($backupUserInput)
	  $backupPassword=GuiCtrlRead($backupPasswordInput)
	  
	  $usmtServer=GuiCtrlRead($usmtServerInput)
	  $usmtUser=GuiCtrlRead($usmtUserInput)
	  $usmtPassword=GuiCtrlRead($usmtPasswordInput)
	  
	  $sysprepServer=GuiCtrlRead($sysprepServerInput)
	  $sysprepUser=GuiCtrlRead($sysprepUserInput)
	  $sysprepPassword=GuiCtrlRead($sysprepPasswordInput)
	  
	  $driversServer=GuiCtrlRead($driversServerInput)
	  $driversUser=GuiCtrlRead($driversUserInput)
	  $driversPassword=GuiCtrlRead($driversPasswordInput)
	  
	  $imageXServer=GuiCtrlRead($imageXServerInput)
	  $imageXUser=GuiCtrlRead($imageXUserInput)
	  $imageXPassword=GuiCtrlRead($imageXPasswordInput)
	  
   EndIf
   IniWrite($sysprepConf,"Servers","useGlobal",$useGlobalServer)
   IniWrite($sysprepConf,"Servers","globalServer",$globalServer)
   IniWrite($sysprepConf,"Servers","globalUser",$globalUser)
   IniWrite($sysprepConf,"Servers","globalPassword",$globalPassword)
   IniWrite($sysprepConf,"Servers","backupServer",$backupServer)
   IniWrite($sysprepConf,"Servers","backupUser",$backupUser)
   IniWrite($sysprepConf,"Servers","backupPassword",$backupPassword)
   IniWrite($sysprepConf,"Servers","usmtServer",$usmtServer)
   IniWrite($sysprepConf,"Servers","usmtUser",$usmtUser)
   IniWrite($sysprepConf,"Servers","usmtPassword",$usmtPassword)
   IniWrite($sysprepConf,"Servers","sysprepServer",$sysprepServer)
   IniWrite($sysprepConf,"Servers","sysprepUser",$sysprepUser)
   IniWrite($sysprepConf,"Servers","sysprepPassword",$sysprepPassword)
   IniWrite($sysprepConf,"Servers","driversServer",$driversServer)
   IniWrite($sysprepConf,"Servers","driversUser",$driversUser)
   IniWrite($sysprepConf,"Servers","driversPassword",$driversPassword)
   IniWrite($sysprepConf,"Servers","imageXServer",$imageXServer)
   IniWrite($sysprepConf,"Servers","imageXUser",$imageXUser)
   IniWrite($sysprepConf,"Servers","imageXPassword",$imageXPassword)
   
   Return 25
EndFunc
Func _Save_Backup()
   $backupFilePath = GUICtrlRead($backupPathInput)
   $backupDrive = GUICtrlRead($backupDriveInput)
   $backupFilename = GUICtrlRead($backupFileNameInput)
   $backupLogFile = GUICtrlRead($backupLogFileInput)
   
   If ($sysprepType='USB') Then
	  $backupDrive = ""
   EndIf
   IniWrite($sysprepConf,"Backup","backupDrive",$backupDrive)
   IniWrite($sysprepConf,"Backup","backupPath",$backupFilePath)
   IniWrite($sysprepConf,"Backup","backupFilename",$backupFilename)
   IniWrite($sysprepConf,"Backup","backupLogFile",$backupLogFile)
   Return 15
EndFunc
Func _Save_USMT()
   $usmtDrive = GUICtrlRead($usmtDriveInput)
   $usmtFilePath = GUICtrlRead($usmtPathInput)
   $usmtTempDir = GUICtrlRead($usmtTempDirInput)
   $usmtXMLFile = GUICtrlRead($usmtXMLFileInput)
   $usmtLogFile = GUICtrlRead($usmtLogFileInput)
   If GuiCtrlRead($usmtListFilesInput) = $GUI_CHECKED Then
	  $usmtListFiles = True
	  $usmtListFilesFileName = GuiCtrlRead($usmtListFilesFileNameInput)
   Else
	  $usmtListFiles = False
	  $usmtListFilesFileName = ""
   EndIf

   If ($sysprepType='USB') Then
	  $usmtDrive = ""
   EndIf
   IniWrite($sysprepConf,"USMT","usmtDrive",$usmtDrive)
   IniWrite($sysprepConf,"USMT","usmtPath",$usmtFilePath)
   IniWrite($sysprepConf,"USMT","usmtTempDir",$usmtTempDir)
   IniWrite($sysprepConf,"USMT","usmtXMLFile",$usmtXMLFile)
   IniWrite($sysprepConf,"USMT","usmtLogFile",$usmtLogFile)
   IniWrite($sysprepConf,"USMT","usmtListFiles",$usmtListFiles)
   IniWrite($sysprepConf,"USMT","listFilesFileName",$usmtListFilesFileName)
   
   Return 20
EndFunc
Func _Save_Drivers()
   $driversDrive = GuiCtrlRead($driversDriveInput)
   $driversFilePath = GuiCtrlRead($driversPathInput)
   
   If ($sysprepType='USB') Then
	  $driversDrive = ""
   EndIf
   IniWrite($sysprepConf,"Drivers","driversDrive",$driversDrive)
   IniWrite($sysprepConf,"Drivers","driversPath",$driversFilePath)
   
   Return 10
EndFunc
Func _Save_ImageX()
   $imageXDrive = GuiCtrlRead($imageXDriveInput)
   $imageXFilePath = GuiCtrlRead($imageXPathInput)
   $defaultWIM =  GuiCtrlRead($defaultWIMInput)
   
   If ($sysprepType='USB') Then
	  $imageXDrive = ""
   EndIf
   IniWrite($sysprepConf,"ImageX","imageXDrive",$imageXDrive)
   IniWrite($sysprepConf,"ImageX","imageXPath",$imageXFilePath)
   IniWrite($sysprepConf,"ImageX","defaultWIM",$defaultWIM)
   
   Return 10
EndFunc
#endregion
#Region CloseScript
Func CloseWindow()
   If $savedConfig=False Then
	  $saveCheck = MsgBox(3,"Save Config?","Do you want to save your current configuration?")
	  If $saveCheck=6 Then
		 _Save_All()
	  EndIf
   EndIf
   If ($fromSysprep=True) Then
	  Run("sysprep.exe")
   EndIf
   Exit
EndFunc
#endregion

;Opens the current configuration file
Func _Open_Config()
   Run(@COMSPEC & ' /c' & ' sysprep_config.ini',@WorkingDir, @SW_HIDE)
EndFunc
;Toggles the disabled and enabled state of the Server Inputs
Func _Server_Input_Toggle()
   If GUICtrlRead($useGlobalCheckBox) = $GUI_CHECKED Then
	  GUICtrlSetState($globalServerInput,$GUI_ENABLE)
	  GUICtrlSetState($globalUserInput,$GUI_ENABLE)
	  GUICtrlSetState($globalPasswordInput,$GUI_ENABLE)
	  
	  GUICtrlSetState($backupServerInput,$GUI_DISABLE)
	  GUICtrlSetState($backupUserInput,$GUI_DISABLE)
	  GUICtrlSetState($backupPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtServerInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtUserInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($sysprepServerInput,$GUI_DISABLE)
	  GUICtrlSetState($sysprepUserInput,$GUI_DISABLE)
	  GUICtrlSetState($sysprepPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($driversServerInput,$GUI_DISABLE)
	  GUICtrlSetState($driversUserInput,$GUI_DISABLE)
	  GUICtrlSetState($driversPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($imageXServerInput,$GUI_DISABLE)
	  GUICtrlSetState($imageXUserInput,$GUI_DISABLE)
	  GUICtrlSetState($imageXPasswordInput,$GUI_DISABLE)
	  $useGlobalServer = true
   Else
	  GUICtrlSetState($globalServerInput,$GUI_DISABLE)
	  GUICtrlSetState($globalUserInput,$GUI_DISABLE)
	  GUICtrlSetState($globalPasswordInput,$GUI_DISABLE)
	  
	  GUICtrlSetState($backupServerInput,$GUI_ENABLE)
	  GUICtrlSetState($backupUserInput,$GUI_ENABLE)
	  GUICtrlSetState($backupPasswordInput,$GUI_ENABLE)
	  GUICtrlSetState($usmtServerInput,$GUI_ENABLE)
	  GUICtrlSetState($usmtUserInput,$GUI_ENABLE)
	  GUICtrlSetState($usmtPasswordInput,$GUI_ENABLE)
	  GUICtrlSetState($sysprepServerInput,$GUI_ENABLE)
	  GUICtrlSetState($sysprepUserInput,$GUI_ENABLE)
	  GUICtrlSetState($sysprepPasswordInput,$GUI_ENABLE)
	  GUICtrlSetState($driversServerInput,$GUI_ENABLE)
	  GUICtrlSetState($driversUserInput,$GUI_ENABLE)
	  GUICtrlSetState($driversPasswordInput,$GUI_ENABLE)
	  GUICtrlSetState($imageXServerInput,$GUI_ENABLE)
	  GUICtrlSetState($imageXUserInput,$GUI_ENABLE)
	  GUICtrlSetState($imageXPasswordInput,$GUI_ENABLE)
	  $useGlobalServer = false
   EndIf
EndFunc
Func _USMT_ListFiles_Input_Toggle()
   If GUICtrlRead($usmtListFilesInput) = $GUI_CHECKED Then
	  GUICtrlSetState($usmtListFilesFileNameInput,$GUI_ENABLE)
   Else
	  GUICtrlSetState($usmtListFilesFileNameInput,$GUI_DISABLE)
   EndIf
EndFunc
Func _Sysprep_Type_Toggle()
   If GuiCtrlRead($sysprepTypeCombo) = 'USB' Then
	  GUICtrlSetState($sysprepUSBDriveInput,$GUI_ENABLE)
	  ;Disable Server Tab
	  GUICtrlSetState($useGlobalCheckBox,$GUI_DISABLE)
	  GUICtrlSetState($globalServerInput,$GUI_DISABLE)
	  GUICtrlSetState($globalUserInput,$GUI_DISABLE)
	  GUICtrlSetState($globalPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($backupServerInput,$GUI_DISABLE)
	  GUICtrlSetState($backupUserInput,$GUI_DISABLE)
	  GUICtrlSetState($backupPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtServerInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtUserInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($sysprepServerInput,$GUI_DISABLE)
	  GUICtrlSetState($sysprepUserInput,$GUI_DISABLE)
	  GUICtrlSetState($sysprepPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($driversServerInput,$GUI_DISABLE)
	  GUICtrlSetState($driversUserInput,$GUI_DISABLE)
	  GUICtrlSetState($driversPasswordInput,$GUI_DISABLE)
	  GUICtrlSetState($imageXServerInput,$GUI_DISABLE)
	  GUICtrlSetState($imageXUserInput,$GUI_DISABLE)
	  GUICtrlSetState($imageXPasswordInput,$GUI_DISABLE)
	  ;Disable Drive Letter boxes
	  GUICtrlSetState($sysprepDriveInput,$GUI_DISABLE)
	  GUICtrlSetState($backupDriveInput,$GUI_DISABLE)
	  GUICtrlSetState($usmtDriveInput,$GUI_DISABLE)
	  GUICtrlSetState($imagexDriveInput,$GUI_DISABLE)
	  GUICtrlSetState($driversDriveInput,$GUI_DISABLE)
   Else
	  GUICtrlSetState($sysprepUSBDriveInput,$GUI_DISABLE)
	  If $useGlobalServer=True Then
		 GUICtrlSetState($useGlobalCheckBox,$GUI_CHECKED)
		 _Server_Input_Toggle()
	  EndIf
	  GUICtrlSetState($useGlobalCheckBox,$GUI_ENABLE)
	  _Server_Input_Toggle()
	  
	  ;Enable Drive Letter boxes
	  GUICtrlSetState($sysprepDriveInput,$GUI_ENABLE)
	  GUICtrlSetState($backupDriveInput,$GUI_ENABLE)
	  GUICtrlSetState($usmtDriveInput,$GUI_ENABLE)
	  GUICtrlSetState($imagexDriveInput,$GUI_ENABLE)
	  GUICtrlSetState($driversDriveInput,$GUI_ENABLE)
   EndIf
EndFunc