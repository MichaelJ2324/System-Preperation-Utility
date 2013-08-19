#include <GuiConstantsEx.au3>
#include <AVIConstants.au3>
#include <File.au3>
#include <Constants.au3>
#include <Array.au3>

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
		 $ping_error=0
	  Else
		 $ping_error = @error
	  EndIf
	  $pingErrors = $ping_error
   Else
	  Local $pingServers[3]
	  For $i = 0 to UBound($serverArray)-1
		 $pingServers[$i]=Ping($serverArray[$i][0])
		 If $pingServers[$i] > 0 Then
			$ping_error=0
		 Else
			$ping_error=@error
			$pingErrors=$pingErrors+1
		 EndIf
	  Next
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
			$mapDriveErrors=$mapDriveErrors+1
			FileWriteLine($batchScriptHandle,"net use " & $drive & " " & $fullPath & " /USER:" & $user & " " & $pw)
		 EndIf
	  EndIf
   Next
   FileClose($batchScriptHandle)
   If $mapDriveErrors > 0 Then
	  If FileExists($mapBatchScript)=1 Then
		 RunWait(@COMSPEC & ' /c ' & $mapBatchScript,@WorkingDir,@SW_HIDE)
		 $mapDriveErrors = 0
		 For $i = 0 to 4
			$drive=$serverArray[$i][2] & ":"
			If FileExists($drive & "\")=0 Then
			   $mapDriveErrors=$mapDriveErrors+1
			EndIf
		 Next
	  Else
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
#region GUI
Func Begin()
   $Main = GUiCreate("Create Data File", 130, 150)
   $fileMenu=GuiCtrlCreateMenu("File")
		   ;Items in File Menu
		   $exitItem=GuiCtrlCreateMenuItem("Exit", $fileMenu)
	   ;Help Menu
	   $helpMenu=GuiCtrlCreateMenu("Help")
		   ;Items 
		   $aboutItem=GUICtrlCreateMenuItem("About",$helpMenu)

	   ;Create Buttons
	   $launchBTN = GUICtrlCreateButton("Go!", 20, 20, 90, 90)
	   
   GUISetState()
   While 1
	   $msg = GUIGetMsg()
	   Select
		   Case $msg = $launchBTN
			   RunWait('"'& @ProgramFilesDir & '\Microsoft Office\Office14\outlook.exe"')
			   LoadState()
			   RunWait(@COMSPEC & ' /c' & "gpupdate /force",@SYSTEMDIR)
			   DriveMapDel('U:')
			   DriveMapDel('W:')
			   If FileExists("C:\Programs.txt") Then
				   Run("C:\Programs.txt")
			   EndIf
			   If FileExists("C:\Printers.txt") Then
				   Run("C:\Printers.txt")
			   EndIf
		   EndSelect
	   If $msg = $GUI_EVENT_CLOSE Then 
		  UnMapDrives()
		  ExitLoop
	   EndIf
   WEnd
EndFunc	
	
#endregion GUI
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
				Exit
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
	 ;$data_error = ;_Interpreter('Data', 1)
   Else
	 ;$data_error = ;_Interpreter('Data', 0)
   EndIf
EndFunc
#endregion Data
Func LoadState()
	If ProcessExists("OUTLOOK.exe") Then
		ProcessClose("OUTLOOK.exe")
	EndIf
	$loadstateCmd = 'loadstate.exe U:\Migration\' & $compName & ' /i:usermigrate.xml /ui:FWMRPC\' & $username & ' /ue:FWMRPC\Administrator /ue:FWMRPC\fwmetals /ue:'& _
					$compName & '\*'
	RunWait(@comspec & ' /k' & $loadstateCMD, 'W:\USMT4.01')
	Run('"'& @ProgramFilesDir & '\Microsoft Office\Office14\OUTLOOK.exe" /importnk2')
EndFunc