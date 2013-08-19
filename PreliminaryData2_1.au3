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
		 EndIf
	  EndIf
   EndIf
EndFunc
#endregion Setup
#region GUICreate
Func Begin()
   $Main = GUiCreate("Create Data File", 300, 140)
   $fileMenu=GuiCtrlCreateMenu("File")
		   ;Items in File Menu
		   $exitItem=GuiCtrlCreateMenuItem("Exit", $fileMenu)
	   ;Help Menu
	   $helpMenu=GuiCtrlCreateMenu("Help")
		   ;Items 
		   $aboutItem=GUICtrlCreateMenuItem("About",$helpMenu)

	   ;Create Buttons
	   $dataBTN = GUICtrlCreateButton("System Info", 15, 20, 80, 80)
	   $programsBTN = GUICtrlCreateButton("Programs List", 110, 20, 80, 80)
	   $printersBTN = GUICtrlCreateButton("Printers List", 205, 20, 80, 80)
	   
   GetServiceTag()
   GuiSetState()
   While 1
	   $msg = GUIGetMsg()
	   Select
		   Case $msg = $dataBTN
			   CreateDataFile()
		   Case $msg = $programsBTN
			   CreateProgramsList()
		   Case $msg = $printersBTN
			   CreatePrinterList()
		   Case $msg = $exitItem
			   $msg = $GUI_EVENT_CLOSE
		   Case $msg = $aboutItem
	   EndSelect
	   If $msg = $GUI_EVENT_CLOSE Then 
		   UnMapDrives()
		   ExitLoop
	   EndIf
   WEnd
EndFunc
#endregion GUI
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
#region Functions
Func GetServiceTag()
	$getServiceTag = Run(@COMSPEC & ' /c' & 'wmic csproduct get identifyingNumber', @SYSTEMDIR, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Sleep(2000)
	$output = StdoutRead($getServiceTag)
	$output = StringStripWS($output, 4) ;STrip of unecessary spaces between words
	$outputarray=StringSplit($output, " ")
	Global $serviceTag=$outputarray[2]
EndFunc
Func CreateDataFile()
	$getData = Run(@COMSPEC & ' /c' & 'wmic computersystem get model, name, username', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Sleep(6000)
	;Parse out data for file
	$test = StdoutRead($getData) ; Read Data from command line output
	$test = StringStripWS($test, 4) ;STrip of unecessary spaces between words
	$testarray=StringSplit($test, " ") ;Split string into array splitting the string by the spaces between words
	$user = StringSplit($testarray[7], "\")
	;Start making GUI
	$verificationForm = GUICreate("Verify Data", 250, 250) 
	GUICtrlCreateLabel("Please Verify all data and change it if necessary.", 10, 10) 
	GUICtrlCreateLabel("Service Tag:", 10, 30) 
	$verifyST = GUICtrlCreateInput($serviceTag, 100, 30, 100)
	GUICtrlCreateLabel("Model:", 10, 60) 
	$verifyModel = GUICtrlCreateInput($testarray[4], 100, 60, 100)
	GUICtrlCreateLabel("Model#:", 10, 90) 
	$verifyModelnum = GUICtrlCreateInput($testarray[5], 100, 90, 100)
	GUICtrlCreateLabel("Comp. Name:", 10, 120) 
	$verifyCompName = GUICtrlCreateInput($testarray[6], 100, 120, 100)
	GUICtrlCreateLabel("User Name:", 10, 150) 
	$verifyUser = GUICtrlCreateInput($user[2], 100, 150, 100)
	$submitBTN = GUICtrlCreateButton("Create Data File", 75, 180, 100, 50)
			
	GUISetState(@SW_Show, $verificationForm)
	While 1
		$verify = GUIGetMsg($verificationForm)
			Select
			Case $verify = $submitBTN
					$serviceTag = GuiCtrlRead($verifyST) 
					$model=GuiCtrlRead($verifyModel)
					$modelnum=GUiCtrlRead($verifyModelnum)
					$compName=GUiCtrlRead($verifyCompName)
					$userName=GUiCtrlRead($verifyUser)
					$keys="Model=" & $model & @LF & "Modelnum=" & $modelnum & @LF & "CompName=" & $compName & @LF & "UserName=" & $userName & @LF
					IniWriteSection($dataFileLocation, $serviceTag, $keys)
					$verify = $GUI_EVENT_CLOSE
					MsgBox(0,"Data File Created","Data added to Data.ini")
			EndSelect
		If $verify = $GUI_EVENT_CLOSE Then 
			GuiSetState(@SW_HIDE, $verificationForm)
			ExitLoop
		EndIf
	WEnd
EndFunc
Func CreateProgramsList()
	;WMIC command to list all programs installed on the current machine
	$getPrograms = Run(@COMSPEC & ' /c' & 'wmic product get name', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	;Open the initial list of programs on a fresh install of windows 7 for comparison with above output
	
	;Prepare inital program list from a fresh Win7 install
	if (FileExists($sysprepDrive & ":\" & $dataIniPath & "\DefaultPrograms.txt")) Then
	   
	   $comparison = FileOpen($sysprepDrive & ":\" & $dataIniPath & "\DefaultPrograms.txt", 0)
	   $arrayNum = _FileCountLines($sysprepDrive & ":\" & $dataIniPath & "\DefaultPrograms.txt")
	
	   Dim $compareArray[$arrayNum]
	   For $x=1 to $arrayNum-1
		   $compareArray[$x-1] = FileReadLine($comparison,$x)
		   If @error=-1 Then
			   ExitLoop
		   Endif
		   $compareArray[$x-1] = StringStripWS($compareArray[$x-1],2)
	   Next
	   FileClose($comparison)
	   ;Progress Bar For creation of data file
	   ProgressOn("Generating List of Programs","Generating List of Programs", "Please wait...")
	   For $x = 1 to 10
		   Sleep(1000)
		   ProgressSet(($x * 10))
	   Next 
	   ProgressSet(100, "Complete")
	   ProgressOff()
	   $programs = StdoutRead($getPrograms) ; Read Data from command line output
	   $programs = StringStripCR($programs) ;STrip of unecessary returns between lines
	   $programsArray = StringSplit($programs, @CRLF)
	   
	   $programFile=FileOpen("C:\Programs.txt", 2)
	   ;Compare programLists to determine which need to be installed
	   $prgs=UBound($programsArray)
	   
	   For $x=0 to $prgs-1
		   $counter=0
		   $programsArray[$x] = StringStripWS($programsArray[$x],2)
		   For $c=0 to $arrayNum-1
			   If StringCompare($programsArray[$x], $compareArray[$c], 0)=0 Then
				   $counter=$counter+1
			   EndIf
		   Next
		   If $counter=0 Then
			   FileWriteLine($programFile, $programsArray[$x])
		   EndIf
	   Next
	   FileClose($programFile)
	   MsgBox(0,"Complete", "The list of programs is completed.")
	Else
	   MsgBox(0,"Error","DefaultPrograms.txt file not found in DataIniPath, " & $sysprepDrive & ":\" & $dataIniPath & "\DefaultPrograms.txt")
	EndIf
EndFunc
Func CreatePrinterList()
	$getPrinter = Run(@COMSPEC & ' /c' & 'wmic printer get name', @SystemDir, @SW_HIDE, $STDERR_CHILD + $STDOUT_CHILD)
	Sleep(2000)
	;Parse out data for file
	$print = StdoutRead($getPrinter) ; Read Data from command line output
	$print = StringStripCR($print) ;STrip of unecessary spaces between words
	$printarray=StringSplit($print, @CRLF) ;Split string into array splitting the string by the spaces between words
	
	$printFile = FileOpen("C:\Printers.txt", 2)
	For $x = 0 to $printarray[0]
		If StringInStr($printarray[$x], "\\") Then
			FileWriteLine($printFile,$printArray[$x])
		ElseIf StringInStr($printarray[$x], "Dymo") Then
			FileWriteLine($printFile,$printArray[$x])
		Elseif StringInStr($printarray[$x], "FaxFinder") Then
			FileWriteLine($printFile,$printArray[$x])
		EndIf
	Next
	FileClose($printFile)
	MsgBox(0,"Complete", "The list of printers is completed.")
EndFunc