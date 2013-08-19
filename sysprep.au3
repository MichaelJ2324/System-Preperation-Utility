#Region Includes
;*****Includes ********
;Includes for Various functions built in to AutoIt.
;Copy.au3 is the reason copying drivers has a progress bar.
#include <GuiConstantsEx.au3>
#include <WindowsConstants.au3>
#include <AVIConstants.au3>
#include <File.au3>
#include <Constants.au3>
#include <Array.au3>
#include <Copy.au3>
#include <ErrorInterpreter.au3>
#include <Date.au3>

#EndRegion Includes

;Initialize program - load global variables and configuration
Init()
;Run setup - Get system information, start logging, check connection and drives
Setup()
;Launch GUI and entire Utility
Begin()

#region Initialize
;Load global variables from config file
Func Init()
   ;Define other global variables
   Global $model, $modelNum, $compName, $platform, $userName, $serviceTag, $dataFileLocation
   Global $g_BackupVerified=False
   ;;Global error varaibles
   Global $usb_error, $data_error, $backup_error, $migration_error, $diskPart_error, $image_error, $bootrec_error, $driver_error, $hotkey_error, $map_drive_error[5]
   ;;Global variables for GUIs (defined gloablly so they can be called from any form at any time)
   Global $Main, $instructionForm, $aboutForm, $filenamePrompt, $filenameBox
   Global $backupBTN, $usmtBTN, $gImageXBTN, $diskPrepBTN, $driversBTN, $bootRecBTN
   
   ;Start pulling global variables from Config File
   Global $sysprepConf="sysprep_config.ini"
   ;Load sysprep Type from config
   Global $sysprepType = IniRead($sysprepConf,"Config","sysprepType","Network")
   Global $sysprepUSBDrive = IniRead($sysprepConf,"Config","usbDriveLabel","")
   ;Load sysprep config from INI
   Global $dataIniFile = IniRead($sysprepConf,"Config","dataFile","data.ini")
   Global $dataIniPath = IniRead($sysprepConf,"Config","dataPath","sysinfo")
   
   ;Load sysprep logging config from INI
   Global $g_logFileDirectory = IniRead($sysprepConf,"Logging","logFileDir","LogFiles")
   Global $g_LogFileName = IniRead($sysprepConf,"Loggin","logFileName","<serviceTag>_sysprepLog.txt")
   Global $logLevel = IniRead($sysprepConf,"Logging","logLevel","5")
   
   ;Load sysprep server configuration from INI
   ;;Check if using global server
   Global $useGlobalServer = IniRead($sysprepConf,"Servers","useGlobal","True")
   If $useGlobalServer=True Then
	  ;;Load all servers as global server 
	  Global $globalServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
	  Global $backupServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
	  Global $usmtServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
	  Global $sysprepServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
	  Global $driversServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
	  Global $imageXServer = IniRead($sysprepConf,"Servers","globalServer","ittest")

	  ;;Load Username and Passwords for servers
	  Global $globalUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
	  Global $globalPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
	  Global $backupUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
	  Global $backupPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
	  Global $usmtUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
	  Global $usmtPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
	  Global $sysprepUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
	  Global $sysprepPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
	  Global $driversUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
	  Global $driversPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
	  Global $imageXUser = IniRead($sysprepConf,"Servers","globalUser","ITTest\sysprep")
	  Global $imageXPassword = IniRead($sysprepConf,"Servers","globalPassword","FWM$sysprep24")
   Else
	  ;;Load all servers seperately	  
	  Global $globalServer = IniRead($sysprepConf,"Servers","globalServer","ittest")
	  Global $backupServer = IniRead($sysprepConf,"Servers","backupStorage","ittest")
	  Global $usmtServer = IniRead($sysprepConf,"Servers","usmtStorage","ittest")
	  Global $sysprepServer = IniRead($sysprepConf,"Servers","sysprepStorage","ittest")
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
   EndIf
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
   Global $usmtLogFile = IniRead($sysprepConf,"USMT","usmtLogFile","<serviceTag>_scanstate.log")
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
   Global $imageXFilePath = IniRead($sysprepConf,"ImageX","imageXPath","images\current")
   Global $imageXDrive = IniRead($sysprepConf,"ImageX","imageXDrive","I")
   Global $defaultWIM = IniRead($sysprepConf,"ImageX","defaultWIM","current.wim")
   
   ;Load Server configuration into Global Server Array
   Global $serverArray[5][5] = [ _
								 [$backupServer,$backupFilePath,$backupDrive,$backupUser,$backupPassword], _
								 [$usmtServer,$usmtFilePath,$usmtDrive,$usmtUser,$usmtPassword], _
								 [$sysprepServer,$sysprepFilePath,$sysprepDrive,$sysprepUser,$sysprepPassword], _
								 [$driversServer,$driversFilePath,$driversDrive,$driversUser,$driversPassword], _
								 [$imageXServer,$imageXFilePath,$imageXDrive,$imageXUser,$imageXPassword] _
							  ]
EndFunc
#endregion Initialize

#region Setup
Func Setup()
   _Init_Interpreter()
   $serviceTag = GetServiceTag()
   If ($sysprepType='USB') Then
	  _Disk()
	  If ($sysprepDrive) Then
		 setAllDrivesForUSB()
		 $dataFileLocation = $sysprepDrive & ":\" & $dataIniPath & "\" & $dataIniFile
	  EndIf
   Else
	  $networkCheck = CheckNetwork()
	  If $networkCheck=0 Then
		 $mappedDrives = MapNetworkDrives()
		 If $mappedDrives = 0 Then
			If FileExists($sysprepDrive & ":\" & $dataIniPath & "\" & $dataIniFile) Then
			   $dataFileLocation = $sysprepDrive & ":\" & $dataIniPath & "\" & $dataIniFile
			Else
			   $dataFileLocation = $sysprepDrive & ":\" & $dataIniPath & "\" & $dataIniFile
			   IniWriteSection($dataFileLocation, $serviceTag, "Model=Test" & @LF & "Modelnum=Test" & @LF & "CompName=Test" & @LF & "Username=Test" & @LF)
			EndIf
			GetData()
		 EndIf
		 BTNswitchOn()
	  Else
		 BTNSwitchOff()
	  EndIf
   EndIf
EndFunc
;Get Service tag of computer
Func GetServiceTag()
   $getServiceTag = Run(@COMSPEC & ' /c' & 'wmic csproduct get identifyingNumber', @SYSTEMDIR, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
   Sleep(3000)
   $output = StdoutRead($getServiceTag)
   $output = StringStripWS($output, 4) ;Strip of unecessary spaces between words
   $outputarray=StringSplit($output, " ")
   Return $outputarray[2]
EndFunc 
Func setAllDrivesForUSB()
   $backupDrive = $sysprepDrive
   $usmtDrive = $sysprepDrive
   $driversDrive = $sysprepDrive
EndFunc
#endregion Setup
#region Network
;Checks for network connection
Func CheckNetwork()
   Local $pingErrors = 0
   If $useGlobalServer=true Then
	  Local $ping = Ping($backupServer)
	  If $ping > 0 Then
		 _Interpreter('Ping',0)
		 $ping_error=0
	  Else
		 $ping_error = @error
		 _Interpreter('Ping',$ping_error)
	  EndIf
	  $pingErrors = $ping_error
   Else
	  Local $pingServers[3]
	  For $i = 0 to UBound($serverArray)-1
		 $pingServers[$i]=Ping($serverArray[$i][0])
		 If $pingServers[$i] > 0 Then
			$ping_error=0
			_Interpreter('Ping',0,$serverArray[$i])
		 Else
			$ping_error=@error
			_Interpreter('Ping',$ping_error,$serverArray[$i])
			$pingErrors=$pingErrors+1
		 EndIf
	  Next
	  If $pingErrors > 0 Then
		 _Interpreter('Ping',5)
	  EndIf
   EndIf
   Return $pingErrors
EndFunc
;Map Network Drives
Func MapNetworkDrives()
   Local $mapDriveErrors,$drive,$fullPath,$user,$pw,$mapBatchScript
   UnMapDrives()
   $mapDriveErrors=0
   $mapBatchScript="mapDrives.bat"
   $batchScriptHandle = FileOpen($mapBatchScript,2)
   For $i = 0 to 4
	  $drive=$serverArray[$i][2] & ":"
	  $fullPath="\\" & $serverArray[$i][0] & "\" & $serverArray[$i][1]
	  $user=$serverArray[$i][3]
	  $pw=$serverArray[$i][4]
	  If FileExists($drive & "\")=0 Then
		 $mapDrive = DriveMapAdd($drive,$fullPath,0,$user,$pw)
		 If $mapDrive = 0 Then
			$map_drive_error[$i]=_Interpreter('Map',@error,$drive)
			$mapDriveErrors=$mapDriveErrors+1
			FileWriteLine($batchScriptHandle,"net use " & $drive & " " & $fullPath & " /USER:" & $user & " " & $pw)
		 Else
			$map_drive_error[$i]=_Interpreter('Map',0,$drive)
		 EndIf
	  Else
		 _Interpreter('Map',3,$drive)
	  EndIf
   Next
   FileClose($batchScriptHandle)
   If $mapDriveErrors > 0 Then
	  _Interpreter('Map',7,$drive)
	  If FileExists($mapBatchScript)=1 Then
		 RunWait(@COMSPEC & ' /c ' & $mapBatchScript,@WorkingDir,@SW_HIDE)
		 $mapDriveErrors = 0
		 For $i = 0 to 4
			$drive=$serverArray[$i][2] & ":"
			If FileExists($drive & "\")=0 Then
			   _Interpreter('Map',10,$drive)
			   $mapDriveErrors=$mapDriveErrors+1
			EndIf
		 Next
	  Else
		 _Interpreter('Map',8)
		 $batchScriptHandle = FileOpen($mapBatchScript,2)
		 For $i = 0 to 4
			$drive=$serverArray[$i][2] & ":"
			$fullPath="\\" & $serverArray[$i][0] & "\" & $serverArray[$i][1]
			$user=$serverArray[$i][3]
			$pw=$serverArray[$i][4]
			If FileExists($drive & "\")=0 Then
			   FileWriteLine($batchScriptHandle,"net use " & $drive & " " & $fullPath & " /USER:" & $user & " " & $pw)
			EndIf
		 Next
		 FileClose($batchScriptHandle)
		 RunWait(@COMSPEC & ' /c ' & $mapBatchScript,@WorkingDir,@SW_HIDE)
		 $mapDriveErrors = 0
		 For $i = 0 to 4
			$drive=$serverArray[$i][2] & ":"
			If FileExists($drive & "\")=0 Then
			   _Interpreter('Map',10,$drive)
			   $mapDriveErrors=$mapDriveErrors+1
			EndIf
		 Next
	  EndIf
   EndIf
   Return $mapDriveErrors
EndFunc
Func UnMapDrives($drive='All')
   If ($drive='All') Then
	   For $i = 0 to 4
		 $drive=$serverArray[$i][2] & ":"
		 DriveMapDel($drive)
	  Next
   Else
	  DriveMapDel($drive&":")
   EndIf
EndFunc
#endregion Network
#Region Removable Disk
;Checks for removeable disk and sets up file locations
Func _Disk()
	;;	Check for disk
	$compDrives=_Find_Disk()
	If (IsArray($compDrives)=1) Then
	  $arraySize=$compDrives[0]
	   For $x=1 to $arraySize-1
		   $drvLabel=DriveGetLabel($compDrives[$x])
		   If $drvlabel="FWMI_SysPrep" Then
			   $sysprepDrive=$compDrives[$x]
		   EndIf
	   Next
	   If $drive="" Then
		   $sysprepDrive = _Choose_Drive($compDrives)
	   EndIf
    Else
	   Exit
    EndIf
EndFunc
Func _Choose_Drive($driveLetters)
	Opt("GUIOnEventMode", 0)
	$arraySize=$driveLetters[0]
	$letterString= "1. " & $driveLetters[1] 
	For $x=2 to $arraySize-1
		$letterString=$letterString & @CRLF & $x & ". " & $driveLetters[$x] & " " & DriveGetLabel($driveLetters[$x])
	Next
	$chooseDrive=GUICreate("Choose Drive:", 400, 210)
	$label20=GuiCtrlCreateLabel("Please choose the disk drive that contains all necessary System Preparation files." & @CRLF & _
								"If you are unsure open a command prompt and check the volumes manually." & @CRLF & _
								"Please choose the number of the drive. Here are the current drives you can choose from:" & @CRLF, 10, 10)
	$label22= GUiCtrlCreateLabel($letterString, 20, 50, 310)
	$label21=GuiCtrlCreateLabel("Choose the number next to the drive letter (i.e. 1 or 2)", 10, 145)
	$cmdBTN=GUICtrlCreateButton("CMD.exe", 340, 60, 50, 30)
	$driveBox=GUICtrlCreateInput("", 10, 160, 150, 20)
	$choose= GUICtrlCreateButton("Submit", 330, 150, 60, 40)
	GUISetState(@SW_SHOW, $chooseDrive)
	While 1
		$dsk=GUIGetMsg($chooseDrive)
		Select
		Case $dsk=$choose
			GUISetState(@SW_HIDE, $chooseDrive)
			$driveNum=GUICtrlRead($driveBox)
			If StringIsInt($driveNum) Then
				$driveNum=Int($driveNum)
				If $driveNum>$arraySize-1 Then
					GUICtrlSetData($driveBox, "")
					GUISetState(@SW_SHOW, $chooseDrive)
					MsgBox(0, "Error!", "Choose a number that is listed.")
				Else
					Opt("GUIOnEventMode", 1)
					Return $driveLetters[$driveNum]
				EndIf
			Else
				GUICtrlSetData($driveBox, "")
				GUISetState(@SW_SHOW, $chooseDrive)
				MsgBox(0, "Error!", "Choose a number that is listed.")
			EndIf
		Case $dsk=$cmdBTN
			_Run_CMD()
		EndSelect	
		If $dsk=$GUI_EVENT_CLOSE Then
			GUISetState(@SW_HIDE, $chooseDrive)
			ExitLoop
		EndIf
	WEnd
	Opt("GUIOnEventMode", 1)
EndFunc
#endregion Removable Disk
#Region GUI
Func begin()
   Opt("GUIOnEventMode", 1)
   ;************************************Graphical User Interface***********************************************************
   ;Create Main Window
   Global $Main = GuiCreate("FWM Sysprep", 750, 550,-1,-1,-1)
   ;
   ;Create Menus in top left corner
   ;;File Menu Contains: 
   ;;;	Ping Check - checks for network connection and reenable buttons if for pingcheck failed.
   ;;;	Open Data File - Opens data file located on flash drive in notepad for editing
   ;;;	Load Data File - Just in case you clicked no at the first prompt, or if you edited the data file you can reload
   ;;;					the variables back into the program.
   ;;;	Exit - Exit
	   $fileMenu=GuiCtrlCreateMenu("File")
		   $dataOpenItem = GuiCtrlCreateMenuItem("Open Data File", $fileMenu)
			   GUICtrlSetOnEvent($dataOpenItem, "_Run_Open_DataFile")
		   $sysPrepLog=GuiCtrlCreateMenuItem("View SysPrep Log", $fileMenu)
			   GUICtrlSetOnEvent($sysPrepLog, "_Log_Sysprep")
			$sysprepConfig = GuiCtrlCreateMenuItem("Configure Sysprep", $fileMenu)
			    GUICtrlSetOnEvent($sysprepConfig, "_Run_ConfigMgr")
		   $exitItem=GuiCtrlCreateMenuItem("Exit", $fileMenu)	
			   GUICtrlSetOnEvent($exitItem, "ExitScript")
   ;;Utilities Menu:
   ;;;	Command Prompt - Opens a blank command prompt for whatever you need or desire
	   $utilsMenu=GuiCtrlCreateMenu("Utilities")
		   $cmdItem=GUICtrlCreateMenuItem("Command Prompt", $utilsMenu)
			   GUICtrlSetOnEvent($cmdItem, "_Run_CMD")
		   $acronisItem=GUICtrlCreateMenuItem("Acronis", $utilsMenu)
			   GUICtrlSetOnEvent($acronisItem, "_Run_Acronis")
   ;;Help Menu:
   ;;;	Instructions - Contains instructions page with tabs to all different functions included in the program
   ;;;	About - Gives version number and such
	   $helpMenu=GuiCtrlCreateMenu("Help")
		   $instructionsItem=GUICtrlCreateMenuItem("Instructions",$helpMenu)
			   GUICtrlSetOnEvent($instructionsItem, "Instructions")
		   $aboutItem=GUICtrlCreateMenuItem("About",$helpMenu)
			   GUICtrlSetOnEvent($aboutItem, "About")
   ;		
   ;**Create Main Tabs to Seperate Windows 7 install process and Windows Xp install process**
	   GUICtrlCreateTab(0,0,802,602)
   ;;*Windows 7 Tab:
   ;;	Contains 6 Buttons - 1 for each process
	   $win7Tab = GUICtrlCreateTabItem("Windows 7")
   ;;;	Image at Top
	   $image1 = GUICtrlCreatePic(".\Pictures\logo.jpg", 0, 40, 750, 315)
   ;;;	Backup Button:
   ;;;;	Right Click Menu has item for viewing the log file.
	   $backupBTN = GUICtrlCreateButton("Backup", 38.5, 400, 80, 80)
		   GUICtrlSetOnEvent($backupBTN, "Backup")
		   $backupRC = GUICtrlCreateContextMenu($backupBTN)
			   $backupLog = GUICtrlCreateMenuItem("View Log", $backupRC)
				   GUICtrlSetOnEvent($backupLog, "_Log_Backup")
   ;;; User State Migration Button:
   ;;;;	Right Click Menu has items for viewing log file, and for viewing file that list files moved.
	   $usmtBTN = GUICtrlCreateButton("USMT", 157, 400, 80, 80)
		   GUICtrlSetOnEvent($usmtBTN, "USMT")
		   $usmtRC = GUICtrlCreateContextMenu($usmtBTN)
			   $usmtLog = GuiCtrlCreateMenuItem("View Log", $usmtRC)
				   GUICtrlSetOnEvent($usmtLog, "_Log_USMT")
			   $usmtFile = GuiCtrlCreateMenuItem("View Files", $usmtRC)
   ;;;	Disk Prep Button: (Disk Part)
   ;;;;	Right Click Menu has item to automate disk part, not fully functional yet
   ;;;;	Also has menu item for viewing log after automation
	   $diskPrepBTN = GUICtrlCreateButton("DiskPart", 275.5, 400, 80, 80)
		   GUICtrlSetOnEvent($diskPrepBTN, "BackupCheck")
		   $diskPrepRC = GUICtrlCreateContextMenu($diskPrepBTN)
			   $diskPart = GuiCtrlCreateMenuItem("Disk Part", $diskPrepRC)
				   GUICtrlSetOnEvent($diskPart, "_Run_DiskPart")
			   $diskPrepLog = GuiCtrlCreateMenuItem("View Log", $diskPrepRC)
				   GUICtrlSetOnEvent($diskPrepLog, "_Log_DiskPrep")
   ;;;	GImageX Button:
   ;;;;	Right Click menu has item to automate process using imagex, not fully functional yet
	   $gImageXBTN = GUICtrlCreateButton("GImageX", 394, 400, 80, 80)
		   GUICtrlSetOnEvent($gImageXBTN, "BackupCheck")
		   $gImageXRC = GUICtrlCreateContextMenu($gImageXBTN)
			   $imagex = GuiCtrlCreateMenuItem("Run ImageX x86", $gImageXRC)
				   GUICtrlSetOnEvent($imagex, "BackupCheck")
   ;;;	BootRec /rebuild button
	   $bootRecBTN = GUICtrlCreateButton("BootRec", 512.5, 400, 80, 80)
		   GUICtrlSetOnEvent($bootRecBTN, "_Run_BootRec")
   ;;;	Drivers button
	   $driversBTN = GUICtrlCreateButton("Drivers", 631, 400, 80, 80)
		   GUICtrlSetOnEvent($driversBTN, "Drivers")
   ;
   ;
   ;;*Windows XP Tab*
   ;;	Contains 4 Buttons - 1 for each Process
	   $winXPTab = GUICtrlCreateTabItem("Windows XP")
   ;;;	Image at Top - soon to be flash create AVI for splash screen and beauty
	   $image2 = GUICtrlCreatePic(".\Pictures\logo.jpg", 0, 40, 750, 315)
   ;;;	Backup Button:
   ;;;;	Right Click Menu has item for viewing the log file.
	   $backupBTNxp = GUICtrlCreateButton("Backup", 100, 400, 80, 80)
		   GUICtrlSetOnEvent($backupBTNxp, "Backup")
		   $backupRCxp = GUICtrlCreateContextMenu($backupBTNxp)
			   $backupLogxp = GUICtrlCreateMenuItem("View Log", $backupRCxp)
				   GUICtrlSetOnEvent($backupLogxp, "_Log_Backup")
   ;;;	Disk Prep Button: (Disk Part)
   ;;;;	Right Click Menu has item to automate disk part, not fully functional yet
   ;;;;	Also has menu item for viewing log after automation
	   $diskPrepBTNxp = GuiCtrlCreateButton("DiskPart", 250, 400, 80, 80)
		   GUICtrlSetOnEvent($diskPrepBTNxp, "DiskPrep")
		   $diskPrepRCxp = GUICtrlCreateContextMenu($diskPrepBTNxp)
			   $diskPartxp = GuiCtrlCreateMenuItem("Disk Part", $diskPrepRCxp)
				   GUICtrlSetOnEvent($diskPartxp, "_Run_DiskPart")
			   $diskPrepLogxp = GuiCtrlCreateMenuItem("View Log", $diskPrepRCxp)
				   GUICtrlSetOnEvent($diskPrepLogxp, "_Log_DiskPrep")
   ;;;	GImageX Button:
   ;;;;	Right Click menu has item to automate process using imagex, not fully functional yet
	   $acronisBTNxp = GuiCtrlCreateButton("Acronis", 400, 400, 80, 80)
   ;;;	BootSect /nt52 Button
	   $bootsectBTN = GuiCtrlCreateButton("BootSect", 550, 400, 80, 80)
   ;;;;;;;;;;;;;;;;;;;;;; End Main User Interface ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
   ;Set Main GUI State
   GUISetOnEvent($GUI_EVENT_CLOSE, "CloseWindow")
   GUISwitch($Main)
   GuiSetState(@SW_SHOW,$Main) 																												
   While 1
	   Sleep(1000)
   WEnd
EndFunc
#endregion

#region Data
Func GetData()
	; Custom Prompt asking if you even need a data file. IF you do, you also have the option of using the current one, if one is found on the usb drive.
	$useDataFilebox = GuiCreate("Data File", 300, 100)
	$usedatafileLabel = GUICtrlCreateLabel("Is this an upgrade?", 15, 15)
	$checkbox1 = GUICtrlCreateCheckbox("Use Data.ini", 115, 11)
	$yesButton = GUICtrlCreateButton("Yes", 160, 60, 50, 25)
	$noButton = GUICtrlCreateButton("No", 225, 60, 50, 25)
	GUISetState(@SW_SHOW, $useDataFilebox)
	GUICtrlSetState($checkbox1, $GUI_CHECKED)
	While 1
		$use = GUIGetMsg($useDataFilebox)
		Select
			Case $use = $yesButton 
				GUISetState(@SW_HIDE, $useDataFilebox)
				If FileExists("C:\Programs.txt") Then
					FileCopy("C:\Programs.txt", "X:\")
				EndIf
				If FileExists("C:\Printers.txt") Then
					FileCopy("C:\Printers.txt", "X:\")
				EndIf
				If GuiCtrlRead($checkbox1) = $GUI_CHECKED Then
					ReadData()
				Else
					GetModel()
					MakeData()
				EndIf
				$use = $GUI_EVENT_CLOSE
			Case $use = $noButton
				GUISetState(@SW_HIDE, $useDataFilebox)
				GetModel()
				$userName = "(username)"
				$compName = "(compName)"
				$use = $GUI_EVENT_CLOSE
				$data_error = _Interpreter('Data', 2)
		EndSelect
		If $use = $GUI_EVENT_CLOSE Then 
			GUISetState(@SW_HIDE, $useDataFilebox)
			ExitLoop
		EndIf
	WEnd
EndFunc
Func GetModel()
   If ($model="") Then
	  $makeAndModel=""
	  $getModel = Run(@COMSPEC & ' /c' & 'wmic computersystem get model', @SYSTEMDIR, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	  Sleep(3000)
	  $output = StdoutRead($getModel)
	  $output = StringStripWS($output, 4) ;STrip of unecessary spaces between words
	  $outputarray=StringSplit($output, " ")
	  For $x = 1 to $outputarray[0]
		 If NOT($outputarray[$x]="Model") Then
			$makeAndModel=$makeAndModel & " " & $outputarray[$x]
			If StringInStr($outputarray[$x], "Latitude1") Then
			   $model = $outputarray[$x]
			   $modelNum = $outputarray[$x + 1]
			   $platform = "Laptop"
			   ExitLoop
			ElseIf StringInStr($outputarray[$x], "Optiplex") Then
			   $model = $outputarray[$x]
			   $modelNum = $outputarray[$x + 1]
			   $platform = "Desktop"
			   ExitLoop
			EndIf
		 EndIf
	  Next
	  If $platform = "" Then
		 PromptForModel($makeAndModel)
	  EndIf
   EndIf
EndFunc
Func PromptForModel($makeAndModel="")
	$modelForm = GuiCreate("SysPrep", 300, 200) 
	$foundLabel = GUICtrlCreateLabel("Found Make And Model:", 50, 10)
	$foundInput = GUICtrlCreateInput($makeAndModel, 25, 25, 250, 25)
	$modelLabel = GUICtrlCreateLabel("Model:", 15, 63)
	$modelInput = GUICtrlCreateInput($model, 100, 60, 125, 25)
	$modelNumLabel = GUICtrlCreateLabel("Model Number:", 15, 93)
	$modelNumInput = GUICtrlCreateInput($modelNum, 100, 90, 125, 25)
	$platformLabel = GUICtrlCreateLabel("Platform:", 15, 123)
	$platformCombo = GUICtrlCreateCombo("Desktop", 100, 120, 125, 25)
	$platformData = GUICtrlSetData($platformCombo,"Laptop",$platform)
	$SubmitBTN = GUICtrlCreateButton("Submit", 100, 150, 100, 30)
	GUISetState(@SW_SHOW,$modelForm)
	While 1
		$info = GUIGetMsg($modelForm)
		Select
			Case $info = $SubmitBTN
			   If NOT($modelInput="" AND $modelNum="" AND $platform="") Then
				  GUISetState(@SW_HIDE,$modelForm)
				  $model=GuiCtrlRead($modelInput)
				  $modelNum=GuiCtrlRead($modelNumInput)
				  $platform=GuiCtrlRead($platformCombo)
				  ExitLoop
			   EndIf
		EndSelect
	WEnd
EndFunc
Func MakeData()
	$infoForm = GuiCreate("SysPrep", 300, 120) 
	$labelComp = GUICtrlCreateLabel("Old Computer Name", 15, 10)
	$compNameBox = GUICtrlCreateInput("", 15, 25, 125, 25)
	$labelUser = GUICtrlCreateLabel("UserName", 155, 10)
	$userNameBox = GUICtrlCreateInput("", 155, 25, 125, 25)
	$SubmitBTN = GUICtrlCreateButton("Submit", 140, 80, 50, 25)
	GUISetState(@SW_SHOW,$infoForm)
	While 1
		$info = GUIGetMsg($infoForm)
		Select
			Case $info = $SubmitBTN
				GUISetState(@SW_HIDE,$infoForm)
				$compName=GuiCtrlRead($compNameBox)
				$userName=GuiCtrlRead($userNameBox)
				If $compName = "" then
					MakeData()
				ElseIf $userName = "" then
					MakeData()
				Else
					$keys="Model=" & $model & @LF & "Modelnum=" & $modelNum & @LF & "CompName=" & $compName & @LF & "UserName=" & $userName & @LF
					IniWriteSection($dataFileLocation, $serviceTag, $keys)
				EndIf
			ExitLoop
		EndSelect
	WEnd
EndFunc
Func ReadData()
   $model=IniRead($dataFileLocation, $serviceTag, "Model", $model)
   $modelNum=IniRead($dataFileLocation, $serviceTag, "Modelnum", $modelNum)
   $compName=IniRead($dataFileLocation, $serviceTag, "CompName", "(computer)")
   $userName=IniRead($dataFileLocation, $serviceTag, "UserName", "(username)")
   If StringInStr($model, "Latitude") Then
	 $platform = "Laptop"
   ElseIf StringInStr($model, "Optiplex") Then
	 $platform = "Desktop"
  Else
	 PromptForModel()
   EndIf
   IF $userName = "newuser" Then
	 $data_error = _Interpreter('Data', 1)
   Else
	 $data_error = _Interpreter('Data', 0)
   EndIf
EndFunc
#endregion Data
#Region Backup
Func Backup()
   ;GUI For Filename Prompt
   Opt("GUIOnEventMode", 1)
   Local $backup = replaceINIVariables($backupFilename)
   If StringInStr($backup,".TIB",2)=0 Then
	  $backup=$backup & ".TIB"
   EndIf
   $filenamePrompt=GUICreate("Backup Filename", 350, 120)
   $label12=GUICtrlCreateLabel("What do you want the backup file to be called. Default naming" & @CRLF & "convention is shown in textbox.", 10, 10)
   $label11=GuiCtrlCreateLabel("Filename:", 10, 37)
   $filenameBox=GUICtrlCreateInput($backup, 10, 50, 250, 20)
   $submitBTN=GUICtrlCreateButton("Submit", 210, 80, 100, 30)
   GUICtrlSetOnEvent($submitBTN, "_Run_Backup")
   GUISetOnEvent($GUI_EVENT_CLOSE, "CloseWindow")
   GUISwitch($filenamePrompt)
   GUISetState(@SW_SHOW, $filenamePrompt)
EndFunc
Func _Run_Backup()
   $backupFN=GUICtrlRead($filenameBox)
   $backupFullPath = $backupDrive & ":\" & $backupFN
   GUISetState(@SW_HIDE, $filenamePrompt)
   $backupLogFullPath = @WorkingDir & "\LogFiles\" & $backupLogFile
   $listPartitions = Run(@COMSPEC & ' /c' & 'TrueImageCmd.exe /list', @ProgramFilesDir & '\Acronis\BackupAndRecovery\', @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
   Sleep(5000)
   $output = StdoutRead($listPartitions)
   $output = StringStripCR($output)
   $outputarray=StringSplit($output, @CRLF)
   Local $c = 0
   Local $line = "Line" 								; = the line that display the C: drive in the /list command by acronis
   Local $partition = ""
   For $x = 0 to $outputarray[0]
	  If StringInStr($outputarray[$x], "C:") Then
		 $line = $outputarray[$x]
		 $newarray = StringSplit($line, " ")
		 $partition = $newarray[1]
		 StringStripWS($partition,8) 
		 $c = $c +1
	  EndIf
   Next
   If $partition = "" Then
	  $backup_error = _Interpreter('Backup', 1)
   ElseIf $partition = "null" Then
	  $backup_error = _Interpreter('Backup', 3)
   ElseIf StringInStr($line, $partition) = 0 Then
	  $backup_error = _Interpreter('Backup', 2)
   Else
	  $AcronisCMD = 'TrueImageCmd.exe /create /partition:' & $partition & ' /progress:on /compression:9 /filename:"' & $backupFullPath & '" /log:"' & $backupLogFullPath & '"'
	  Run(@COMSPEC & ' /k' & $AcronisCMD, @ProgramFilesDir & '\Acronis\BackupAndRecovery\')
	  $backup_error = _Interpreter('Backup', 0)
   EndIf
EndFunc
#endregion Backup
#Region User State Migration
Func USMT()
   $usmtTempDir = replaceINIVariables($usmtTempDir)
   $tempPath = $usmtDrive & ":\" & $usmtTempDir
   $migrationFolder = $usmtDrive & ':\Migration\' & $compName
   $usmtLogFile = replaceINIVariables($usmtLogFile)
   $usmtLogFolder = @WorkingDir & "\LogFiles\" & $usmtLogFile
   If $usmtListFiles=True Then
	  $usmtListFilesFileName = replaceINIVariables($usmtListFilesFileName)
	  $usmtListFilesCmd = '/listfiles:"' & @WorkingDir & '\LogFiles\' & $usmtListFilesFileName & '"'
   Else
	  $usmtListFilesFolder = ''
   EndIf
   $setWorkDirCMD = 'SET USMT_WORKING_DIR=' & $tempPath
   $scanstateCMD = 'scanstate.exe ' & $migrationFolder & ' /offlineWinDir:C:\Windows /i:' & $usmtXMLFile & ' /ui:FWMRPC\' & $userName & ' /ue:' & $compName & '\* /ue:FWMRPC\Administrator /ue:FWMRPC\fwmetals /v:13 /L:"' & $usmtLogFolder & '" ' & $usmtListFilesCmd & ' /c /o'
   Run(@COMSPEC & ' /k' & $setWorkDirCMD & "&&" & $scanstateCMD, @ProgramFilesDir & '\USMT4.01\')
   $migration_error = _Interpreter('USMT', 0)
EndFunc
#EndRegion USMT
#Region Disk Prepartion
Func BackupCheck()
   If ($g_BackupVerified=False) Then
	  $backupFilename = replaceINIVariables($backupFilename)
	  If StringInStr($backupFilename,".TIB",2)=0 Then
		 $backupFilename=$backupFilename & ".TIB"
	  EndIf
	  $backupFullName = $backupDrive & ":\" & $backupFilename
	  
	  Opt("GUIOnEventMode", 0)
	  $backupCheck=GUICreate("Backup Check", 350, 100)
	  $label10=GuiCtrlCreateLabel("Backup File:", 10, 10)
	  $backupFileBox=GUICtrlCreateInput($backupFullName, 10, 25, 250, 20)
	  $browseBTN= GUICtrlCreateButton("Browse", 265, 23, 50, 25)
	  $submitBTN=GUICtrlCreateButton("Verify Backup", 210, 55, 100, 30)
	  $pathcheck=0
	  GUISetState(@SW_SHOW,$backupCheck)
	  While 1
		$checkObj = GUIGetMsg($backupCheck)
		Select
			Case $checkObj=$browseBTN
			   $chosenBackup = FileOpenDialog("Choose Backup File",$backupDrive & ":\","",1,$backupFilename)
			   If @error Then
				  _Interpreter('BackupCheck',@error,$chosenBackup)
			   Else
				  GUICtrlSetData($backupFileBox, $chosenBackup)
			   EndIf
			Case $checkObj=$submitBTN
				$chosenBackup=GUICtrlRead($backupFileBox)
				$checkObj = $GUI_EVENT_CLOSE
		EndSelect
		If $checkObj = $GUI_EVENT_CLOSE Then
			GUISetState(@SW_HIDE,$backupCheck)
			ExitLoop
		EndIf
	  WEnd
	  Opt("GUIOnEventMode", 1)
	  If FileExists($chosenBackup) Then
		$size = FileGetSize($chosenBackup)
		If ($size/1048576) < 3072 Then
			_Interpreter('BackupCheck',1,$chosenBackup)
			$check = MsgBox(4, "Warning: Backup", "Your backup file is quite small. This can be caused by failed backup due to interruption or corruption. It is highly recommended that you rerun the backup process, however you may continue at your own risk. Continue?")
			If $check = 6 Then
			   _Interpreter('BackupCheck',4)
				If @GUI_CTRLID = $diskPrepBTN Then
					_Run_AutomatedDiskPrep()
				ElseIf @GUI_CTRLID = $GImageXBTN Then
					_Run_GImageX()
				ElseIf @GUI_CTRLID = $imagex Then
					_Run_ImageX()
				EndIf
			Endif
		 Else
			_Interpreter('BackupCheck',0,$chosenBackup)
			$backup_check = MsgBox(4, "System: Valid Backup", "Your backup file appears to be valid. Are you sure you want to continue?")
			If $backup_check = 6 Then
			   _Interpreter('BackupCheck',6)
				If @GUI_CTRLID = $diskPrepBTN Then
					_Run_AutomatedDiskPrep()
				ElseIf @GUI_CTRLID = $GImageXBTN Then
					_Run_GImageX()
				ElseIf @GUI_CTRLID = $imagex Then
					_Run_ImageX()
				EndIf
			Else
				MsgBox(0, "Abort!", "Remember to backup all data before wiping the partition.")
			EndIf
		EndIf
	 Else
		 _Interpreter('BackupCheck',1,$chosenBackup)
		 $check= MsgBox(4, "System: Missing Backup", "No backup file found. Please run the backup process to ensure no loss of data. However you may continue at your risk. Continue?")
		 If $check = 6 Then
			_Interpreter('BackupCheck',6)
			 If @GUI_CTRLID = $diskPrepBTN Then
				 _Run_AutomatedDiskPrep()
			 ElseIf @GUI_CTRLID = $GImageXBTN Then
				 _Run_GImageX()
			 ElseIf @GUI_CTRLID = $imagex Then
				 _Run_ImageX()
			 EndIf
		 Else
			 MsgBox(0, "Abort!", "Run backup before proceeding!")
		 EndIf
	  EndIf
	  $g_BackupVerified = True
   Else
	  If @GUI_CTRLID = $diskPrepBTN Then
		   _Run_AutomatedDiskPrep()
	  ElseIf @GUI_CTRLID = $GImageXBTN Then
		   _Run_GImageX()
	  ElseIf @GUI_CTRLID = $imagex Then
		   _Run_ImageX()
	  EndIf
   EndIf
EndFunc

Func _Find_Disk($disk="null")
	If $disk="null" Then
		$listPartitions = Run(@COMSPEC & ' /c' & 'TrueImageCmd.exe /list', @ProgramFilesDir & '\Acronis\BackupAndRecovery\', @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
		Sleep(6000)
		$output = StdoutRead($listPartitions)
		If $output = "" Then
			$diskPart_error = _Interpreter('DiskPart', 1)
			MsgBox(0, "Error: Operation Aborted!", "No Disks found. Please run System Preparation Utility again.")
			Return $diskPart_error
		Else
			$output = StringStripCR($output)
			$outputarray=StringSplit($output, @CRLF)
			$num=0
			For $x = 0 to $outputarray[0]
				If StringInStr($outputarray[$x], ":") Then
					$num=$num+1
				EndIf
			Next
			Dim $lines[$num]
			$c=0
			For $x = 0 to $outputarray[0]
				If StringInStr($outputarray[$x], ":") Then
					$lines[$c]=$outputarray[$x]
					$c=$c+1
				EndIf
			Next
			$c=0
			For $x =0 to $num-1
				If Not(StringInStr($lines[$x], "Disk")) Then
					$lines[$x]=StringStripWS($lines[$x], 4)
					$tempArray=StringSplit($lines[$x], " ")
					For $a=0 to $tempArray[0]
						If StringInStr($tempArray[$a], ":") Then
							$tempArray[$a]=StringTrimLeft($tempArray[$a], 1)
							$tempArray[$a]=StringTrimRight($tempArray[$a], 1)
							$lines[$x]=$tempArray[$a]
						EndIf
					Next
					$c=$c+1
				EndIf
			Next
			Dim $driveLetters[$c+1]
			$driveLetters[0]=$c+1
			$c=1
			For $x =0 to $num-1
				If Not(StringInStr($lines[$x], "Disk")) Then
					$driveLetters[$c]=StringStripWS($lines[$x], 8)
					$c=$c+1
				EndIf
			Next
			Return $driveLetters
		EndIf
	Else
		If NOT(DriveGetLabel($disk)="FWMI_SysPrep") Then
			$listPartitions = Run(@COMSPEC & ' /c' & 'TrueImageCmd.exe /list', @ProgramFilesDir & '\Acronis\BackupAndRecovery\', @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
			Sleep(6000)
			$output = StdoutRead($listPartitions)
			If $output = "" Then
				$diskPart_error = _Interpreter('DiskPart', 1)
				MsgBox(0, "Error: Operation Aborted!", "Automation Error, operation aborted. View Log File for more information.")
				Return $diskPart_error
			Else
				$output = StringStripCR($output)
				$outputarray=StringSplit($output, @CRLF)
				$line = "Line" 
				$partitionAmount = 0
				$diskNumber = 1
				For $x = 0 to $outputarray[0]
					If StringInStr($outputarray[$x], $disk) Then
							$line = $outputarray[$x]
							$diskNumber = StringLeft($line, 1)
					EndIf
				Next
				If $line = "Line" or $line = "" Then
					$diskPart_error = _Interpreter('DiskPart', 2)
					MsgBox(0, "Error: Aborted!", "Automation Error, operation aborted. View Log File for more information.")
				Else
					For $x = 0 to $outputarray[0]
						If StringInStr($outputarray[$x], $diskNumber & "-") Then
							$partitionAmount = $partitionAmount + 1
						EndIf
					Next
				EndIf
				Local $partitionInfo[3]
				$partitionInfo[0]=$line
				$partitionInfo[1]=$diskNumber
				$partitionInfo[2]=$partitionAmount
				Return $partitionInfo
			EndIf
		Else
			Return 666
		Endif
	EndIf
 EndFunc
 
Func _Run_AutomatedDiskPrep()
	$listPartitions = Run(@COMSPEC & ' /c' & 'TrueImageCmd.exe /list', @ProgramFilesDir & '\Acronis\BackupAndRecovery\', @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Sleep(6000)
	$output = StdoutRead($listPartitions)
	If $output = "" Then
		$diskPart_error = _Interpreter('DiskPart', 1)
		MsgBox(0, "Error: Operation Aborted!", "Automation Error, operation aborted. View Log File for more information.")
	Else
		$output = StringStripCR($output)
		$outputarray=StringSplit($output, @CRLF)
		$line = "Line" 
		$partitionAmount = 0
		$diskNumber = 1
		For $x = 0 to $outputarray[0]
			If StringInStr($outputarray[$x], "C:") Then
					$line = $outputarray[$x]
					$diskNumber = StringLeft($line, 1)
			EndIf
		Next
		If $line = "" Then
			$diskPart_error = _Interpreter('DiskPart', 2)
			MsgBox(0, "Error: Aborted!", "Automation Error, operation aborted. View Log File for more information.")
		Else
			For $x = 0 to $outputarray[0]
				If StringInStr($outputarray[$x], $diskNumber & "-") Then
					$partitionAmount = $partitionAmount + 1
				EndIf
			Next
			$AutoFile = FileOpen(@SystemDir & "\auto.txt", 2)
			$error = 0
			If $partitionAmount = 0 Then
			FileWriteLine($AutoFile, "Select Disk 0")
			Else
				For $x = 1 to $partitionAmount
					If StringInStr($line, $diskNumber & "-" & $x) Then
						FileWriteLine($AutoFile, "Select Disk " & ($diskNumber - 1))
						FileWriteLine($AutoFile, "Select Partition " & $x)
						FileWriteLine($AutoFile, "Delete Partition")
						If $partitionAmount > 2 Then
							For $i = 2 to $partitionAmount
								If StringInStr($line, $diskNumber & "-" & $i) Then
									;Do Nothing
								Else
									FileWriteLine($AutoFile, "Select Partition " & $i)
									FileWriteLine($AutoFile, "Delete Partition")
								EndIf
							Next
						EndIf
						$error = 0
					Else
						$error = $error + 1
						If $error = $partitionAmount Then
							$error = 3
						Else
							$error = 0
						EndIf
					EndIf
				Next
			EndIf
		EndIf
		IF $error = 0 Then
			FileWriteLine($AutoFile, "Create Partition Primary")
			FileWriteLine($AutoFile, "Format FS=NTFS Quick")
			FileWriteLine($AutoFile, "Active")
			FileWriteLine($AutoFile, "Assign letter=C")
			FileClose($AutoFile)
			Runwait(@COMSPEC & ' /c' & 'diskpart /s ' & @SystemDir & '\auto.txt > ' & @SystemDir & 'log.txt', @SystemDir)
			$diskPart_error = _Interpreter('DiskPart', 0)
			MsgBox(0, "Success!", "Disk Part successfully prepared the drive.")
		Else
			$diskPart_error = _Interpreter('DiskPart', 3)
		EndIf
	EndIf
EndFunc
Func _Run_DiskPart()
	Run(@COMSPEC & ' /c' & 'diskpart', @SystemDir)
	_Interpreter('DiskPart', 0)
EndFunc
#endregion Disk Preparation
#Region ImageX
Func _Run_GImageX()
	Run('GImageX86.exe')
	$image_error = _Interpreter('Image', 0)
EndFunc
Func _Run_ImageX()
   If $defaultWIM="" Then
	  MsgBox(4096,"No Default Wim!","No Default WIM file is set in configuration. Please change the ini file and reload script to use ImageX automation.")
	  _Run_GImageX()
   Else
	  If StringInStr($defaultWIM,".wim",2)=0 Then
		 $defaultWIM = $defaultWIM & ".wim"
	  EndIf
	  $imageFile = '"' & $imageXDrive & ':\Win7\Current\' & $defaultWIM
	  Run(@comspec & ' /k imagex.exe /apply "' & $imageFile & '" 1 C:\')
	  $image_error = _Interpreter('Image', 0)
   EndIf
EndFunc
#endregion ImageX
#region Boot Record
Func _Run_BootRec()
	Run(@COMSPEC & ' /k' & 'bootrec /rebuildbcd', @SystemDir)
	$bootRec_error = _Interpreter('Boot', 0)
EndFunc
Func _Run_BootSect()	
	Run(@COMSPEC & ' /k' & 'bootsect /nt52 sys', @SystemDir)
	$bootRec_error = _Interpreter('Boot', 1)
EndFunc
#endregion Boot Record
#Region Drivers
Func Drivers()
	$defaultsrcPath=$driversDrive & ':\' & $platform & '\' & $model & ' ' & $modelnum & '\x86\Win7\DriverPack'
	$defaultdestPath = 'C:\drivers\computer'
	;Prompt to make sure directory is correct
	$pathPrompt=GUICreate("Driver Paths", 350, 125)
	$label9=GuiCtrlCreateLabel("Source Path:", 10, 7)
	$srcPathBox=GUICtrlCreateInput($defaultsrcPath, 10, 20, 250, 20)
	$browseSrcBTN= GUICtrlCreateButton("Browse", 270, 20, 40, 20)
	$label10=GuiCtrlCreateLabel("Destination Path:", 10, 50)
	$destPathBox=GUICtrlCreateInput($defaultdestPath, 10, 63, 250, 20)
	$browseDestBTN= GUICtrlCreateButton("Browse", 270, 63, 40, 20)
	$submitBTN=GUICtrlCreateButton("Copy Drivers", 210, 85, 100, 30)
	$driversPath=$defaultsrcPath
	$destPath=$defaultdestPath
	$pathcheck=0
	GUISetState(@SW_SHOW,$pathPrompt)
	While 1
		$pathObj = GUIGetMsg($pathPrompt)
		Select
			Case $pathObj=$browseSrcBTN
				$newsrcpath = FileSelectFolder("Choose Driver Folder", "", "4", "Y:\")
				If NOT($newsrcpath="") Then
					GUICtrlSetData($srcPathBox, $newsrcpath)
				EndIf
			Case $pathObj=$browseDestBTN
				$newdestpath = FileSelectFolder("Choose Driver Folder", "", "4", "C:\")
				If NOT($newdestpath="") Then
					GUICtrlSetData($destPathBox, $newdestpath)
				EndIf
			Case $pathObj=$submitBTN
				$driversPath=GUICtrlRead($srcPathBox)
				$destPath=GuiCtrlRead($destPathBox)
				$pathcheck=1
				$pathObj = $GUI_EVENT_CLOSE
		EndSelect
		If $pathObj = $GUI_EVENT_CLOSE Then
			GUISetState(@SW_HIDE,$pathPrompt)
			ExitLoop
		EndIf
	WEnd
	
	If $pathcheck>0 Then
		DirCreate($destPath)
		If Not _Copy_OpenDll() Then
			MsgBox(16, '', 'DLL not found.')
			$driver_error = _Interpreter('Driver', 1)
		Else
			$destPathSize = DirGetSize($destPath)
			If $destPathSize > 0 Then
				$driverSize = DirGetSize($driversPath)
				If $destPathsize < $driverSize Then
					$internal_dir_error = 4;Drivers not copied correctly (greater or smaller)
				Else
					$internal_dir_error = 5 ;Drivers already copied
				EndIf
			EndIf
			If $driversStatus = 0 Then
				CopyDrivers($driversPath, $destPath)
				$driver_error = _Interpreter('Driver', 0)
				Return 1
			ElseIf $driversStatus = 1 Then
				If $internal_dir_error = 4 Then
					DirRemove($destPath)
					DirCreate($destPath)
					CopyDrivers($driversPath, $destPath)
					$driver_error = _Interpreter('Driver', 0)
				EndIf
			EndIf
		EndIf
	EndIf
EndFunc
Func CopyDrivers($source,$dest)
	Local $Copy = False
	$progressBox = GuiCreate("Copying Files...", 400, 150)
	$fromLabel = GUICtrlCreateLabel("Copying Files from " & $source & " ...", 20, 30, 360, 20)
	$Progress = GUICtrlCreateProgress(20, 50, 360, 20)
	$percentComplete = GUICtrlCreateLabel("0% Complete", 20, 75, 360, 20)
	GUISetState(@SW_SHOW, $progressBox)
	_Copy_CopyDir($source,$dest)
	$Copy = 1
	While 1
		$msg2 = GUIGetMsg($progressBox)
		If $Copy Then
			$State = _Copy_GetState()
			If $State[0] Then
				$Data = Round($State[1] / $State[2] * 100)
				If GUICtrlRead($Progress) <> $Data Then
					GUICtrlSetData($Progress, $Data)
					GuiCtrlSetData($percentComplete, $Data & "% Complete")
				EndIf
			Else
				Switch $State[5]
					Case 0
						GUICtrlSetData($Progress, 100)
						MsgBox(64, '', 'Files were successfully copied.', 0, $progressBox)
						$driversStatus = 0
						$msg2 = $GUI_EVENT_CLOSE
					Case 1235 ; ERROR_REQUEST_ABORTED
						MsgBox(16, '', 'File copying was aborted.', 0, $progressBox)
						$driversStatus = 1235
						$msg2 = $GUI_EVENT_CLOSE
					Case Else
						MsgBox(16, '', 'File was not copied.' & @CR & @CR & $State[5], 0, $progressBox)
						$driversStatus = 500
						$msg2 = $GUI_EVENT_CLOSE
				EndSwitch
				GUICtrlSetData($Progress, 0)
				GuiCtrlSetData($percentComplete, $Data & "% Complete")
				$Copy = 0
			EndIf
		EndIf
		Select
			Case $msg2 = $GUI_EVENT_CLOSE
				GUISetState(@SW_HIDE, $progressBox)
				_Copy_Abort()
				ExitLoop
		EndSelect
	WEnd
	Return $driversStatus
EndFunc
#Endregion Drivers
#Region Random Small Functions
Func _Run_CMD()
   Run("cmd.exe", @SystemDir)
EndFunc
Func _Run_Acronis()
   Run(@ProgramFilesDir & "\Acronis\BackupAndRecovery\TrueImage.exe")
EndFunc
Func _Run_Open_DataFile()
   Run("notepad.exe " & $sysprepDrive & ":\" & $dataIniPath & "\" & $dataIniFile)
EndFunc
Func BTNswitchOff()
	  GUICtrlSetState($backupBTN,$GUI_DISABLE)
	  GUICtrlSetState($usmtBTN,$GUI_DISABLE)
	  GUICtrlSetState($gImageXBTN,$GUI_DISABLE)
	  GUICtrlSetState($diskPrepBTN,$GUI_DISABLE)
	  GUICtrlSetState($driversBTN,$GUI_DISABLE)
	  GUICtrlSetState($bootRecBTN,$GUI_DISABLE)
EndFunc
Func BTNswitchOn()
	  GUICtrlSetState($backupBTN, $GUI_ENABLE)
	  GUICtrlSetState($usmtBTN, $GUI_ENABLE)
	  GUICtrlSetState($gImageXBTN, $GUI_ENABLE)
	  GUICtrlSetState($diskPrepBTN, $GUI_ENABLE)
	  GUICtrlSetState($driversBTN, $GUI_ENABLE)
	  GUICtrlSetState($bootRecBTN, $GUI_ENABLE)
EndFunc
Func HotKeyPress()
	  GUICtrlSetState($backupBTN, $GUI_ENABLE)
	  GUICtrlSetState($usmtBTN, $GUI_ENABLE)
	  GUICtrlSetState($gImageXBTN, $GUI_ENABLE)
	  GUICtrlSetState($diskPrepBTN, $GUI_ENABLE)
	  GUICtrlSetState($driversBTN, $GUI_ENABLE)
	  GUICtrlSetState($bootRecBTN, $GUI_ENABLE)
	  $hotkey_error = _Interpreter('Hotkey', 1)
EndFunc
#endregion
#Region View Log Files
Func _Log_Sysprep()
	Run(@COMSPEC & ' /c' & $currentLog, @WorkingDir, @SW_HIDE)
EndFunc
Func _Log_Backup()
   $backupLogFile = replaceINIVariables($usmtLogFile)
   $backupLogFolder = @WorkingDir & "\LogFiles\" & $backupLogFile
	Run(@COMSPEC & ' /c' & $backupLogFolder, @WorkingDir, @SW_HIDE)
EndFunc
Func _Log_USMT()
   $usmtLogFile = replaceINIVariables($usmtLogFile)
   $usmtLogFolder = @WorkingDir & "\LogFiles\" & $usmtLogFile
	Run(@COMSPEC & ' /c' & $usmtLogFolder, @WorkingDir, @SW_HIDE)
 EndFunc	
Func _Run_ConfigMgr()
   Run('sysprepConfigMgr.exe sysprep "True"',@WorkingDir)
   Exit
EndFunc
#endregion
#Region Instruction Page
Func Instructions()
	Opt("GUIOnEventMode", 1)
	$instructionForm = GuiCreate("FWM SysPrep - Instructions", 700, 500)
		GUICtrlCreateTab(0,0,702, 502)
		$programTab = GUICtrlCreateTabItem("Sysprep")
			$title1 = GUICtrlCreateLabel("Sysprep Process", 7, 30)
			$description1 = GuiCtrlCreateLabel("Windows 7 Process:" & @CRLF & _ 																				; Can only come this far
											   "	1. Backup - (For Upgrades from XP) Acronis TruImageCMD.exe. Backs up the hard drive in a viewable format. " & @CRLF & _ 
											   "	2. USMT - (For Upgrades from XP) User State migration Tool. Copies all personal files in user directories plus " & @CRLF & _ 
											   "	   .pst files and .nk2 files. After upgrade USMT copies these files to coresponding directories." & @CRLF & _ 
											   "	3. DiskPart - Windows command line utility for formatting and partitioning hard drive." & @CRLF & _ 
											   "	4. GImageX - GUI version of ImageX used to extract the Windows 7 WIM file onto the hard drive." & @CRLF & _ 
											   "	5. BootRec - The /rebuildbcd parameter rebuilds the boot manager on the hard drive to point toward our new" & @CRLF & _ 
											   "	   image on the C: drive." & @CRLF & _
											   "	6. Drivers - Copies drivers from Fileserv2 to specified area on new Windows install." & @CRLF & _
											   " "  & @CRLF & _
											   "Windows XP Process:"  & @CRLF & _
											   "	1. Backup - Same as above"  & @CRLF & _
											   "	2. Used to reformat and partition hard drive for new install of XP."  & @CRLF & _
											   "	3. GImageX - Used to extract XP image to computer." & @CRLF & _
											   " ", 7, 50, 683, 500)
		
		$programTab2 = GUICtrlCreateTabItem("Sysprep Con't")
			$title2 = GuiCtrlCreateLabel("Things to Know", 7, 30)
			$description2 = GuiCtrlCreateLabel("FAQs:" & @CRLF & _
											   "Question: My buttons are disabled how to I re-enable them?"  & @CRLF & _
											   "Answer: If your buttons are disabled it is because you do not have a connection to Fileser2. To re-enable" & @CRLF & _
											   "	 the buttons you have to either extablish a connection to Fileserv2 and re-run the pingcheck or you" & @CRLF & _
											   "	can use the HotKey (Ctrl + Alt + M) to manually override your network connection status and enable" & @CRLF & _
											   "	the buttons. I suggest establishing a connection since most utilities push or pull data from Fileserv2"  & @CRLF & _
											   "	however, if you are using this jsut as a utility to run commands go right ahead and override them." & @CRLF & _
											   " " & @CRLF & _
											   "Question: What is the Data.txt file? And what data resides in it?" & @CRLF & _
											   "Answer: The data.txt file is located on the root of the thumbdrive you are using. This file is used to store" & @CRLF & _
											   "	information about the current system such as the model type and number. The format of the text file is" & @CRLF & _
											   "	the following." & @CRLF & _
											   "		1. Model - Such as Optiplex or Latitude" & @CRLF & _
											   "		2. Model Number - Such as 990, 780, gx520, or E6400"  & @CRLF & _
											   "		3. ComputerName - (On new installs defualt is NEWPC1)" & @CRLF & _
											   "		4. UserName - (On new installs default is newuser)"  & @CRLF & _
											   "	These lines are loaded into variables and put into different commands throughout the script. The data file" & @CRLF & _
											   "	could have been made when you ran PreSysPrep.exe on an upgrade machine, or manually when running this " & @CRLF & _
											   "	script. If you forgot to run the PreSysPrep.exe script on an upgrade machine this script will allow " & @CRLF & _
											   "	you to add the computername and username to the data file."  & @CRLF & _
											   " ", 10, 50, 683, 500)
		$backupTab = GUICtrlCreateTabItem("Backup")
			$title3 = GUICtrlCreateLabel("Backup Process", 10, 30)
			$description3 = GuiCtrlCreateLabel("Uses TrueImageCMD.exe, a command line utility provided by Acronis."  & @CRLF & _
											   " "  & @CRLF & _
											   "	After clicking the button you are prompted for the partition number in #-# notation. The input box should list"  & @CRLF & _
											   "the current drive number - partition number that the C: drive is on. You should verify that this number is the same" & @CRLF & _
											   "as the number in the input box. If it is not the same please change the number to the correct number in the correct" & @CRLF & _
											   "#-# notation as noted." & @CRLF & _
											   "	If the inputbox does not show the current drive number - partition number of the C: drive then you will need to" & @CRLF & _
											   "verify it manually. To do this open up an command prompt and navigate to " & @ProgramFilesDir & "\Acronis\BackupAndRecovery\" & @CRLF & _ 
											   "directory. Once you are in this directory run the following command to view the current drives and parition on the " & @CRLF & _
											   "machine and verify which one is the C: drive. After you have submitted the correct drive the a command window will open"  & @CRLF & _
											   "and begin the backup process. If the command fails for any reason you can manually enter it using the following command:" & @CRLF & _
											   " " & @CRLF & _
											   " ", 10, 50, 683)
			$command1 = GUICtrlCreateInput('TrueImageCmd.exe /create /partition:(#-#) /progress:on /compression:9 /filename:"U:\'& $userName & ' ' & $modelnum & ' ' & @MDAY & ' ' & @MON & ' ' & @YEAR & '.TIB" /log:' & @SystemDir & '\LogFiles\backup.log', 10, 250, 680, 20)
		$usmtTab = GUICtrlCreateTabItem("USMT")
			$title4 = GUICtrlCreateLabel("USMT Process", 10, 30)
			$description4 = GUICtrlCreateLabel("Uses Scanstate.exe in the User State Migration Tool 4.01, provided by Microsoft." & @CRLF & _
											   " " & @CRLF & _
											   "	After Clicking the button you are prompted to enter in a name for the temporary working directory. The User State Migration Tool" & @CRLF & _
											   "needs a working directory for temporary files while it is migrating data from the computer. You can leave the default if you like but"& @CRLF & _
											   "sometimes it fails, so don't be afraid to switch it up. Note this folder does get deleted right after the process complete so do not" & @CRLF & _
											   "name it after the computer name or user name." & @CRLF & _
											   "	If the command fails you can manually enter it. To do this you must be in the " & @ProgramFilesDir & "\USMT4.01\ directory. After" & @CRLF & _
											   "you are in this directory you must set the temporary working directory using the first command shown below. After you have set the" & @CRLF & _
											   "working directory you can then run the Scanstate command listed below."& @CRLF & _
											   "	Whether you manually entered it or it ran right away the command will generate a log file you can view as well as a file with a list "& @CRLF & _
											   "of the files it migrated. Notice these files will not remain on the thumbdrive and will be deleted after you exit. So if you want to keep" & @CRLF & _
											   "the files pull up a command prompt and use xcopy or copy to copy the files to your destination. I may add a button to do it later." & @CRLF & _
											   " "& @CRLF & _
											   " ",10, 50, 683)
			$command2 = GUICtrlCreateInput("SET USMT_WORKING_DIR=" & $usmtDrive & ":\" & $usmtTempDir, 10, 250, 683, 20)
			$command3 = GUICtrlCreateInput('scanstate.exe ' & $usmtDrive & ':\Migration\' & $compName & ' /offlineWinDir:C:\Windows /i:usermigrate.xml /ui:FWMRPC\' & $userName & ' /ue:' & $compName & '\* /ue:FWMRPC\Administrator /ue:FWMRPC\fwmetals /v:13 /L:' & @SystemDir & '\LogFiles\scanstate.log /listfiles:' & @SystemDir & '\LogFiles\Files.txt  /c /o', 10, 280, 683, 20)
		$diskPartTab = GUICtrlCreateTabItem("DiskPart")
			$title5 = GuiCtrlCreateLabel("DiskPart Process", 10, 30)
			$description5 = GUICtrlCreateLabel("Uses Disk Part utility by Microsoft." & @CRLF & _
											   " " & @CRLF & _
											   "	After clicking the button Disk Part is launched. For upgrades you will delete the primary partition and then create a new primary partition." & @CRLF & _
											   " For a new computer delete all partitions except the Dell OEM partition. These are the commands you need to know to create, delete, and partition a drive." & @CRLF & _ 
											   "		1. List Disk - list the disks curently on the machine." & @CRLF & _
											   "		2. Select Disk - Selects the disk you want to work on." & @CRLF & _
											   "		3. List Partition - list the partitions on the currently selected disk." & @CRLF & _
											   "		4. Select Partition - selects the parition you want to work on. " & @CRLF & _
											   "		5. Delete Partition - deletes the selected partition." & @CRLF & _
											   "		6. Create Partition Primary - creates a primary partition. And selects this partition to work on." & @CRLF & _
											   "		7. Format FS=NTFS Quick - formats the selected partition to NTFS using the quick way. Do NOT forget the QUICK command." & @CRLF & _
											   "		8. Active - marks the partition as active." & @CRLF & _
											   "		9. Assign (Optional: Letter = C) - Assign a drive letter. Will need to be drive C so if it does not set to C by default you can" & @CRLF & _
											   "		   change the drive letter using the letter command." & @CRLF & _
											   "		10. List Volume - list the volumes on the machine, which shows drive letter, size, name, and partition format." & @CRLF & _
											   "	Optional Commands:" & @CRLF & _
											   "		Clean - cleans the drive of all partitions. Only use on drives that do not have the Dell OEM partition." & @CRLF & _
											   " 	You can also right click and hit Automate DiskPart and it does the process for you. However if you have already clicked the" & @CRLF & _
											   "	normal DiskPart button it won't allow you to because DiskPart has been ran once already." & @CRLF & _
											   " " & @CRLF & _
											   " ", 10, 45, 683)
		$gimageXTab = GuiCtrlCreateTabItem("GImageX")
			$title6 = GuiCtrlCreateLabel("GImageX86 Process", 10, 30)
			$description6 = GUICtrlCreateLabel("Uses GImageX86.exe, the graphical utility to ImageX." & @CRLF & _
											   " " & @CRLF & _
											   "	After clicking the button GImageX86 is launched. In GImageX go to the apply tab. In this tab select the source directory using the browse"  & @CRLF & _
											   "button and select W:\Win7\ or W:\WinXP\ and then the proper image. Then select the destination directory using the browse button and click" & @CRLF & _ 
											   "on the C:\ drive. After these two things are set you can select apply and wait until the image is applied. The following is a screenshot of "  & @CRLF & _
											   "GImageX with the correct directories." & @CRLF & _ 
											   " " & @CRLF & _
											   " ", 10, 45, 683)
			$image3 = GUICtrlCreatePic("gimagex.jpg", 25, 200, 538, 169) 
		$bootRecTab = GUICtrlCreateTabItem("BootRec")
			$title7 = GuiCtrlCreateLabel("BootRec Process", 10, 30)
			$description7 = GuiCtrlCreateLabel("Uses BootRec /rebuildbcd command line utility to rebuild the boot manager for Windows 7." & @CRLF & _
											   " "  & @CRLF & _
											   "	After clicking the button the command BootRec /rebuildbcd is launched in the command prompt. Verify that it recognized the proper" & @CRLF & _
											   " Windows partition and type Y or Yes to confirm. It will then compete the process and you can exit."  & @CRLF & _
											   " "  & @CRLF & _
											   " ", 10, 45, 683)
		$driversTab = GUICtrlCreateTabItem("Drivers")
			$title8 = GUICtrlCreateLabel("Process for Copying Drivers", 10, 30)
			$description8 = GuiCtrlCreateLabel("Uses AutoIT built in Copy Function, with custom DLL. "  & @CRLF & _
											   " " & @CRLF & _
											   "	After clicking the button a progress window will open and show you the percentage of drivers copied. If it failes use XCOPY commands below." & @CRLF & _  
											   " " & @CRLF & _
											   " ", 10, 45, 683)
											   
			$command4 = GUICtrlCreateInput('xcopy "' & $driversDrive & ':\' & $platform & '\' & $model & ' ' & $modelnum & '\x86\Win7\DriverPack" C:\drivers\computer\ /S /Y', 10, 160, 683, 20)
			$command4 = GUICtrlCreateInput('xcopy "' & $driversDrive & ':\DriverPacks\x86" C:\drivers\general\ /S /Y', 10, 190, 683, 20)
;;	Set Instruction Page state to shown
	GUISetOnEvent($GUI_EVENT_CLOSE, "CloseWindow")
	GUISwitch($instructionForm)
	GUISetState(@SW_SHOW,$instructionForm)
EndFunc
#endregion Instruction Page
#Region About
;About form
Func About()
	Opt("GUIOnEventMode", 1)
	Global $aboutForm = GUICreate("FWM Sysprep - About", 400, 300)
	GuiCtrlCreateLabel("Version 6.0 " & @CRLF & "Made By: Mike Russell",10,10)
	GUISetOnEvent($GUI_EVENT_CLOSE, "CloseWindow")
	GUISwitch($aboutForm)
	GUISetState(@SW_SHOW, $aboutForm)
EndFunc
#EndRegion About
#Region Exits
Func CloseWindow()
   If @GUI_WINHANDLE = $Main Then
	  $g_LogFileDirectory=replaceINIVariables($g_LogFileDirectory)
	  $g_LogFileName=replaceINIVariables($g_LogFileName)
	  $sysprepFilePath=replaceINIVariables($sysprepFilePath)
	  $newLogFileDir = $sysprepDrive & ":\" & $sysprepFilePath & "\" & $g_LogFileDirectory & "\" & $g_LogFileName
	  $moveLog = _Move_Log($newLogFileDir)
	  If ($moveLog = 1) Then
		 _Close_Log()
		UnMapDrives()
		 EXIT
	  Else
		 _Interpreter("Log File", "Failed to move log file.")
		 MsgBox(0,"Error!","Log file failed to copy to " & @CRLF & $newLogFileDir)
		 Exit
	  EndIf
   ElseIf @GUI_WINHANDLE = $instructionForm Then
	  GUISetState(@SW_HIDE, $instructionForm)
	  GUISwitch($Main)
   ElseIf @GUI_WinHandle = $aboutForm Then
	  GUISetState(@SW_HIDE, $aboutForm)
	  GUISwitch($Main)
   Elseif @GUI_WINHANDLE = $filenamePrompt Then
	  GUISetState(@SW_HIDE, $filenamePrompt)
	  GUISwitch($Main)
   EndIf
EndFunc
#endregion
#Region String Functions
Func replaceINIVariables($iniString)
   Local $variables[10] = ["<serviceTag>","<model>","<modelNum>","<platform>","<userName>","<compName>","<day>","<mon>","<year>","<rand>"]
   For $i=0 to UBound($variables)-1
	  Switch $variables[$i]
		 Case "<serviceTag>"
			If NOT($serviceTag="") Then
			   $iniString=StringReplace($iniString,$variables[$i],$serviceTag)
			EndIf
		 Case "<model>"
			If NOT($model="") Then
			   $iniString=StringReplace($iniString,$variables[$i],$model)
			EndIf
		 Case "<modelNum>"
			If NOT($modelNum="") Then
			   $iniString=StringReplace($iniString,$variables[$i],$modelNum)
			EndIf
		Case "<platform>"
			If NOT($platform="") Then
			   $iniString=StringReplace($iniString,$variables[$i],$platform)
			EndIf
		 Case "<userName>"
			If NOT($userName="") Then
			   $iniString=StringReplace($iniString,$variables[$i],$userName)
			EndIf
		 Case "<compName>"
			If NOT($compName="") Then
			   $iniString=StringReplace($iniString,$variables[$i],$compName)
			EndIf
		 Case "<day>"
			$iniString=StringReplace($iniString,$variables[$i],@MDAY)
		 Case "<mon>"
			$iniString=StringReplace($iniString,$variables[$i],@MON)
		 Case "<year>"
			$iniString=StringReplace($iniString,$variables[$i],@YEAR)
		 Case "<rand>"
			$random = Random(1111, 9999)
			$iniString=StringReplace($iniString,$variables[$i],$random)
	  EndSwitch
   Next
   Return $iniString
EndFunc
#endregion