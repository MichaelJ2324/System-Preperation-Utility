;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;Error Interpreter for SySprep;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Func _Init_Interpreter($logFileName = "sysprepLog.txt", $logFileDirectory="LogFiles", $newLog=True, $logLevel = 5)
   ;Create Log File and put title and date on it.	
   Local $initialLogFile=$logFileName
   Global $currentLog, $logFile
   If FileExists($logFileDirectory & "\" & $initialLogFile) Then
	  If $newLog=True Then
		 $currentLog = $logFileDirectory & "\" & $initialLogFile
		 $logFile = FileOpen($currentLog,2)
		 _Write_Header()
	  Else
		 $currentLog = $logFileDirectory & "\" & $initialLogFile
	  EndIf
   Else
	  $currentLog = $logFileDirectory & "\" & $initialLogFile
	  $logFile = FileOpen($currentLog,8)
	  _Write_Header()
   EndIf
EndFunc
;;
Func _Write_Header()
   FileWriteLine($logFile, "Fort Wayne Metals - System Preparation Utitlity")
   FileWriteLine($logFile, @Mon & "/" & @MDAY & "/" & @YEAR & " @ " & _NowTime())
   FileClose($logFile)
EndFunc
;;
Func _Move_Log($toLocation)
   $logFileMovement = FileMove ($currentLog, $toLocation, 9)
   If $logFileMovement=1 Then
	  $currentLog = $toLocation
   EndIf
   Return $logFileMovement
EndFunc
;;
Func _Close_Log()
   $logFile = FileOpen($currentLog,1)
   FileWriteLine($logFile, "Script end: " & _NowTime())
   FileClose($logFile)
   MsgBox(0,"Logs Copied","Log file has been copied to " & @CRLF & $currentLog) 
EndFunc
;
Func _Interpreter($function, $error, $extra = 'null')
   $logFile = FileOpen($currentLog, 1)
   Switch $function
	  Case 'Ping'
		 _Ping_Errors($error, $extra)
	  Case 'Map'
		 _MapDrive_Errors($error, $extra)
	  Case 'USB'
		 _USB_Errors($error, $extra)
	  Case 'Data'
		 _Data_Errors($error)
	  Case 'Backup'
		 _Backup_Errors($error)
	  Case 'BackupCheck'
		 _BackupCheck_Errors($error,$extra)
	  Case 'USMT'
		 _USMT_Errors($error)
	  Case 'Diskpart'
		 _DiskPart_Errors($error)
	  Case 'Image'
		 _Image_Errors($error)
	  Case 'Boot'
		 _BootRec_Errors($error)
	  Case 'Driver'
		 _Driver_Errors($error)
	  Case 'Hotkey'
		 _Hotkey_Errors($error)
	  Case Else
		 _Other_Errors($error,$extra)
   EndSwitch
   FileClose($logFile)
EndFunc
#Region Ping Errors
;When the function fails (returns 0) @error contains extended information:
;; 1 = Host is offline
;; 2 = Host is unreachable
;; 3 = Bad destination
;; 4 = Other errors
;; 5 = Specific Server Error
Func _Ping_Errors($error,$server = "null" )
	If $error = 0 Then
		FileWriteLine($logFile, "Network Connection Established.")
	Elseif $error = 1 Then
		FileWriteLine($logFile, "Network connection failed. Host appears to be offline.")
	Elseif $error = 2 Then
		FileWriteLine($logFile, "Network connection failed. Host is unreachable.")
	ElseIf $error = 3 Then
		FileWriteline($logFile, "Network connection filed. Destination does not exist.")
	Elseif $error = 4 Then
		FileWriteLine($logFile, "Network connection failed. Unknown error occured. Contact your system administrator for help.")
	EndIf
	Return $error
EndFunc
#endregion
#Region Map Drive Errors
;When the function fails (returns 0) @error contains extended information:
;; 1 = Undefined / Other error. @extended set with Windows API return
;; 2 = Access to the remote share was denied
;; 3 = The device is already assigned
;; 4 = Invalid device name
;; 5 = Invalid remote share
;; 6 = Invalid password
Func _MapDrive_Errors($error, $drive = "null")
   Select
	  Case $error = 0 
		 FileWriteLine($logFile, $drive & " drive successfully mapped.")
	  Case $error = 1 
		 FileWriteLine($logFile, $drive & " drive did not map correctly. Try mapping drive using 'NET USE' command to see error.")
	  Case $error = 2 
		 FileWriteLine($logFile, $drive & " drive did not map correctly. Access to the remote shared was denied.")
	  Case $error = 3
		 FileWriteLine($logFile, $drive & " drive did not map correctly. The device is already assigned.")
	  Case $error = 4 
		 FileWriteLine($logFile, $drive & " drive did not map correctly. Invalid device name.")
	  case $error = 5 
		 FileWriteLine($logFile, $drive & " drive did not map correctly. Invalid remote share.")
	  Case $error = 6
		 FileWriteLine($logFile, $drive & " drive did not map correctly. Invalid password.")
	  Case $error = 7
		 FileWriteLine($logFile, "Not all drives mapped, retrying as batch script.")
	  Case $error = 8
		 FileWriteLine($logFile, "Map.bat file not found. Creating new batch file.")
	  Case $error = 9
		 FileWriteLine($logFile, "File creation failed. Drives not mapped.")
	  Case $error = 10
		 FileWriteLine($logFile, "Map retry failed for " & $drive)
   EndSelect
   Return $error
EndFunc
#endregion
#Region USB Errors
Func _USB_Errors($error, $drive = "null")
	Select
	Case $error = 0 
		FileWriteLine($logFile, "Found drive " & $drive & ".")
	Case $error = 1 
		FileWriteLine($logFile, "No removeable drive found. Setting variables to defaults.")
	EndSelect
	Return $error
EndFunc
#endregion
#Region Data Errors
Func _Data_Errors($error)
   Select
	  Case $error = 0
		 FileWriteLine($logFile, "System make, model, computername, and username loaded successfully with specific data." & @CRLF & _
								"Make: " & $model & @CRLF & _ 
								"Model: " & $modelNum & @CRLF & _ 
								"Computername: " & $compName & @CRLF & _ 
								"Username: " & $userName)
	  Case $error = 1
		 FileWriteLine($logFile, "System make, model, computername, and username loaded successfully with default NEW BUILD data.")
	  Case $error = 2
		 FileWriteLine($logFile, "System make, model, computername, and username loaded successfully with default data.")
	  Case $error = 3
		 FileWriteLine($logFile, "Programs.txt file found. Moved to local storage.")
	  Case $error = 4
		 FileWriteLine($logFile, "Programs.txt file not found.")
	  Case $error = 5
		 FileWriteLine($logFile, "Printers.txt file found. Moved to local storage.")
	  Case $error = 6
		 FileWriteLine($logFile, "Printers.txt file not found.")
	  Case $error = 9
		 FileWriteLine($logFile, "Data.txt found on removeable media.")
	  Case $error = 99
		 FileWriteLine($logFile, "Data.txt not found on removeable media.")
   EndSelect
EndFunc
#endregion
#Region HotKey
Func _Hotkey_Errors($error)
	Select
	Case $error = 1
		FileWriteLine($logFile, "Manual Override key used. I hope you know what you are doing.")
	EndSelect
EndFunc
#endregion
#Region Backup Errors
Func _Backup_Errors($error)
	Select
	Case $error = 0
		FileWriteLine($logFile, "Backup operation ran. View backup log in " & @SystemDir & "\Logfiles\backup.log")
	Case $error = 1
		FileWriteLine($logFile, "No partition selected to backup. Exiting. Please try again and enter in a partition number.")
	Case $error = 2
		FileWriteLine($logFile, "Chosen partition does not match the results found by the system.")
	Case $error = 3
		FileWriteLine($logFile, "Sysprep did not find C: drive on computer. Run command manually to backup partition.")
	EndSelect
 EndFunc
 Func _BackupCheck_Errors($error,$chosenFile="")
	  Select
		 Case $error=0
			FileWriteLine($logFile,"Valid backup file was chosen: " & $chosenFile)
		 Case $error=1
			FileWriteLine($logFile,"Could not open file, " & $chosenFile & ".")
		 Case $error=4
			FileWriteLine($logFile,"Chosen backup file: " & $chosenFile & ", was smaller than 3GBs")
		 Case $error=5
			FileWriteLine($logFile,"Missing backup file.")
		 Case $error=6
			FileWriteLine($logFile,"Procedded after backup check warning.")
	  EndSelect
 EndFunc
#endregion
#Region USMT Errors
Func _USMT_Errors($error)
	Select
	Case $error = 0
		FileWriteLine($logFile, "USMT operation ran. View migration log in " & @SystemDir & "\Logfiles\scanstate.log. Files.txt will also show you which files it moved.")
	Case Else
		FileWriteLine($logFile, "USMT operation ran. Unknown Error Occured")
	EndSelect
EndFunc
#endregion
#Region DiskPart Errors
Func _DiskPart_Errors($error)
	Select
	Case $error = 0
		FileWriteLine($logFile, "Disk Part operation Ran. Drive is now blank and has been formatted.")
	Case $error = 1
		FileWriteLine($logFile, "Automation Error: Could not read command line output, operation aborted.Run Disk Part Manually.")
	Case $error = 2
		FileWriteLine($logFile, "Automation Error: Could not find C: drive, operation aborted. Run Disk Part Manually.")
	Case $error = 3
		FileWriteLine($logFile, "Automation Error: Unknown configuration occurred, drive not formatted. Run Disk Part Manually.")
	EndSelect
EndFunc
#endregion
#Region ImageX Errors
Func _Image_Errors($error)
	Select
	Case $error = 0
		FileWriteLine($logFile, "Drive Imaged successfully.")
	EndSelect
EndFunc
#endregion
#Region BootRec Errors
Func _BootRec_Errors($error)
   Select
   Case $error=0
	  FileWriteLine($logFile,"Bootrec /rebuildbcd was ran.")
   Case $error=1
	  FileWriteLine($logFile,"Bootsect /nt52 was ran for System drives.")
   EndSelect
EndFunc
#endregion
#Region Driver Errors
Func _Driver_Errors($error)
	Select
	Case $error = 0
		FileWriteLine($logFile, "Drivers copied successfully.")
	Case $error = 1
		FileWriteLine($logFile, "Copy.DLL did not load or was not found.")
	Case $error = 2
		FileWriteLine($logFile, "Drivers copied successfully.")
	EndSelect
EndFunc
#endregion
#Region Random Errors
Func _Other_Errors($error,$log)
   FileWriteLine($logFile, $error & " " & $log)
EndFunc