#$language = "VBScript"
#$interface = "1.0"
'----------------------------------------------------------------------------------
'	KIDSSHIELD SCRIPT. SSH TO JUNIPER SSG FW. 
'----------------------------------------------------------------------------------
	Dim strDirectoryLCL
strDirectoryLCL = "C:\Users\All Users\Vandyke"
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const MAX_LEN = 130
Const CRT_REG_INSTALL = "HKEY_LOCAL_MACHINE\SOFTWARE\VanDyke\SecureCRT\Install\Main Directory"
Const CRT_REG_SESSION = "HKEY_CURRENT_USER\Software\VanDyke\SecureCRT\Config Path"
Dim nResult
Dim strLine
Dim nOverwrite
Dim strMonthMaxFileName, strFileString, strSkip, strFileButton, strFileInventory, strFileSession, strCRTexe
Dim strDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke, strLogFile
Dim strDeviceID, strAccountID
Dim nDebug, ShowDebug
Dim nIndex, nInd, nCount, newIndex
Dim objDebug, objSession, objFSO, objEnvar, objButtonFile, objShell
Dim vSession
Dim nStartHH, nEndHH, n, i, nRetries
Dim strUserProfile, vLine, strScreenUser
Dim nCommand, vCommand, vCmdSrx, strSessionSRX, strSession
Dim strAction, strTab
Dim vInventory, nInventory
strCRTexe = "\SecureCRT.exe"""
strFileInventory = "inventory.ini"
strFileSession = "sessions.txt"
strDirectory = "\\MEDIA\_PublicFolder\KidsSchild\"
strDirectoryWork = "C:\KidsSchild\DVLP"
strDirectoryUpdate = "\\HIGS\Install_My\Tools_Networking\KidsSchild"
nDebug = 1
ShowDebug = False
strVersion = "None"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("WScript.Shell")
' Set objEnvar = CreateObject("WScript.Shell")
Sub Main()
'----------------------------------------------------------------
'	Open log File
'----------------------------------------------------------------
			n = 5
			i = 0
			nRetries = 5
				Do While i < nRetries
					On Error Resume Next
					Err.Clear
					    strLogFile = strDirectoryLCL & "\Log\" & "debug-terminal-f" & i & ".log"
						If objFSO.FileExists(strLogFile) Then 
						    Set objFile = objFSO.GetFile(strLogFile)
							If objFile.Size > 1000000 Then 
							    Set objFile = Nothing
								objFSO.DeleteFile strLogFile, True
							End If
						End If 
						Set objDebug = objFSO.OpenTextFile(strLogFile,ForAppending,True)
						Select Case Err.Number
							Case 0
								Exit Do
							Case 70
								i =  i + 1
								n = 3
								crt.sleep 100 * n
							Case Else 
								Exit Do		
						End Select
				Loop
				On Error goto 0
	'------------------------------------------------------------------
	'	CHECK NUMBER OF ARGUMENTS AND EXIT IF LESS THEN 3
	'------------------------------------------------------------------
	If crt.Arguments.Count < 3 Then
		If IsObject(objDebug) Then
			Call TrDebug("CRT: WRONG NUMBER OF ARGUMENTS", "ERROR", objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug("CRT: Use Arguments: <strDevice>,  <End Hours>, <End Minutes>", "", objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug("CRT: New Arfuments format:  to read cli commands from file: ff|ssg|srx|combo|closefw,  file-name, Tab-Name", "", objDebug, MAX_LEN, 1, nDebug)
			objDebug.close
		End If
		crt.quit
		Exit Sub
	End If
	stdOutFile = strDirectoryLCL & "\Temp\" & crt.Arguments(1)
	strTab = crt.Arguments(2)
	Call TrDebug ("CRT:  BEGIN TELNET SCRIPT: " & strTab , "", objDebug, MAX_LEN, 3, nDebug)	
	'------------------------------------------------------------------
	'	LOAD TELNET SESSION CONFIGURATION
	'------------------------------------------------------------------
    strSession = "Null" : strSessionSRX = "Null"
	If objFSO.FileExists(strDirectoryLCL & "\" & strFileSession) Then 
		nSession = GetFileLineCountSelect(strDirectoryLCL & "\" & strFileSession, vSession,"#","NULL","NULL", 0)
		strFolder = Split(vSession(2), ",")(0)
		strSessionSRX = Split(vSession(2), ",")(1)		
        If UBound(Split(vSession(2), ",")) >= 2 Then strSession = Split(vSession(2), ",")(2)
		Call TrDebug ("CRT:  SESSION FOLDER: " & strFolder , "", objDebug, MAX_LEN, 1, nDebug)
		Call TrDebug ("CRT:  SESSION 1: " & strSessionSRX , "", objDebug, MAX_LEN, 1, nDebug)
		Call TrDebug ("CRT:  SESSION 2: " & strSession , "", objDebug, MAX_LEN, 1, nDebug)		
		strDirectory = vSession(1)
		If Right(strDirectory,1) = "\" Then 
			strDirectory = Left(strDirectory,Len(strDirectory) - 1)
			If IsObject(objDebug) and nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": Remote Server Folder: " & strDirectory End If
		End If
        If nSession > 0 Then	strDirectoryVandyke = vSession(0)	End If	' - Vandyke folder to run SecureCRT from 
		If nSession > 3 Then	strDirectoryWork = vSession(3)		End If	' - Work directory scripts are installed to 
		If nSession > 4 Then    strDirectoryUpdate = vSession(4)	End If	' - Source directory to take updates from 
		If nSession > 5 Then    strVersion = vSession(5)			End If	' - Current version of the Package/Launcer
	End If	
	'------------------------------------------------------------
	'   Start Monitoring Log file
	'------------------------------------------------------------
	strBinFolder = strDirectoryWork & "\Bin"
	If ShowDebug Then 
	    Set objBat = objFSO.OpenTextFile(Replace(strLogFile,"log","bat"),ForWriting,True)
		objBat.WriteLine strBinFolder & "\tail.exe -f " &  """" & strLogFile & """"
		objBat.close
		Set objBat = Nothing  
		objShell.run """" & Replace(strLogFile,"log","bat") & """",1
	End If
	'-----------------------------------------------------------------
	'  CHECK SECURECRT IS INSTALLED ON THE SYSTEM AND UP TO DATE
	'-----------------------------------------------------------------
	Dim vCRTupdate
	Redim vCRTupdate(1)
	Dim strCRT_ConfigFolder, SecureCRT_Installed
	Dim strCRT_SessionFolder
	Dim strCRT_HostKeyFolder
	Dim SSH2File, vSSH2, nSSH2, SSH2Exists, RestartRequired
	Dim objHostKeyFolder, colFiles
	RestartRequired = 0
	On Error Resume Next
		SecureCRT_Installed = True
		Err.Clear
		strCRT_InstallFolder = objShell.RegRead(CRT_REG_INSTALL)
		if Err.Number <> 0 Then 
			SecureCRT_Installed = False
			strCRT_InstallFolder = vSession(0)
		End If
		If Right(strCRT_InstallFolder,1) = "\" Then strCRT_InstallFolder = Left(strCRT_InstallFolder,Len(strCRT_InstallFolder)-1)
		Err.Clear
		strCRT_ConfigFolder = objShell.RegRead(CRT_REG_SESSION)
		if Err.Number <> 0 Then 
			strCRT_ConfigFolder = "C:"
			Call TrDebug ("CRT: CAN'T GET SecureCRT Session Folder", "ERROR", objDebug, MAX_LEN, 1, 1)
			Exit Sub
		Else
			If Right(strCRT_ConfigFolder,1) = "\" Then strCRT_ConfigFolder = Left(strCRT_ConfigFolder,Len(strCRT_ConfigFolder)-1)
			strCRT_SessionFolder = strCRT_ConfigFolder & "\Sessions"
			Call TrDebug("CRT: Session Folder: " & strCRT_SessionFolder, "", objDebug, MAX_LEN, 1, 1)		
		End If
	On Error Goto 0
	strDirectoryVandyke = strCRT_InstallFolder
	strCRTexe = """" & strDirectoryVandyke & strCRTexe	
	SSH2File = strCRT_ConfigFolder & "\SSH2.ini"
	Do
	    If GetFileLineCountSelect(strDirectoryLCL & "\Temp\"  & "crtupdate.ini", vCRTupdate,"NULL","NULL","NULL",0) > 0 Then 
		    If GetObjectLineNumber( vCRTupdate, UBound(vCRTupdate), GetScreenUserCrt(), False) > 0 Then 
				Call TrDebug ("CRT: CONFIGURATION IS UP TO DATE: " , "OK", objDebug, MAX_LEN, 1, 1)
			    Exit Do				
			End If
		End If 	   
    	'-----------------------------------------------------
		'   GET LOCATION OF THE HOST KEYs
		'-----------------------------------------------------
		If Not objFSO.FileExists(SSH2File) Then 
			Call TrDebug ("CRT: SSH2.ini DOESN'T EXISTS ","" , objDebug, MAX_LEN, 1, 1)				
			strCRT_HostKeyFolder = strCRT_ConfigFolder & "\Known Hosts"
			Call WriteStrToFile(SSH2File, "S:""Host Key Database Location""=" & strCRT_HostKeyFolder & "\", 1, 1, 0)
			RestartRequired = 1
			If Not objFSO.FolderExists(strCRT_HostKeyFolder) Then 
				Set objHostKeyFolder = objFSO.CreateFolder(strCRT_HostKeyFolder)
				Call TrDebug ("CRT: NEW HOSTKEY FOLDER: " & strCRT_HostKeyFolder,"CREATED", objDebug, MAX_LEN, 1, 1)				
				Set objHostKeyFolder = Nothing
			End If
		Else
			Call GetFileLineCountSelect(SSH2File, vSSH2,"NULL","NULL","NULL",0)
			nSSH2 = GetObjectLineNumber( vSSH2, UBound(vSSH2), "Host Key Database Location", False) 
			If nSSH2 = 0 Then 
				Call TrDebug ("CRT: CAN'T FIND REFERENCE TO HOST KEY FOLDER IN SSH2.ini","" , objDebug, MAX_LEN, 1, 1)				
				strCRT_HostKeyFolder = strCRT_ConfigFolder & "\Known Hosts"
				Call WriteStrToFile(SSH2File, "S:""Host Key Database Location""=" & strCRT_HostKeyFolder & "\", 1, 1, 0)
				Call TrDebug ("CRT: NEW RECORD FOR HOST KEY FOLDER CREATED IN SSH2.ini: ","", objDebug, MAX_LEN, 1, 1)				
				Call TrDebug ("CRT: ","S:""Host Key Database Location""=" & strCRT_HostKeyFolder & "\", objDebug, MAX_LEN, 1, 1)				
				If Not objFSO.FolderExists(strCRT_HostKeyFolder) Then 
					Set objHostKeyFolder = objFSO.CreateFolder(strCRT_HostKeyFolder)
					Call TrDebug ("CRT: NEW HOSTKEY FOLDER: " & strCRT_HostKeyFolder,"CREATED", objDebug, MAX_LEN, 1, 1)								
					Set objHostKeyFolder = Nothing
				End If
			Else
				strCRT_HostKeyFolder = Split(vSSH2(nSSH2-1),"=")(1)
				If Len(strCRT_HostKeyFolder) < 2 Then 
					strCRT_HostKeyFolder = strCRT_ConfigFolder & "\Known Hosts"
					Call WriteStrToFile(SSH2File, "S:""Host Key Database Location""=" & strCRT_HostKeyFolder & "\", nSSH2, 1, 0)
					RestartRequired = 2
					If Not objFSO.FolderExists(strCRT_HostKeyFolder) Then 
						Set objHostKeyFolder = objFSO.CreateFolder(strCRT_HostKeyFolder)
						Set objHostKeyFolder = Nothing
					End If
				Else 
					If Right(strCRT_HostKeyFolder,1) = "\" Then strCRT_HostKeyFolder = Left(strCRT_HostKeyFolder,Len(strCRT_HostKeyFolder)-1)
					Call TrDebug ("CRT: FOUND RECORD FOR HOST KEY FOLDER: " & strCRT_HostKeyFolder,"", objDebug, MAX_LEN, 1, 1)								
					If Not objFSO.FolderExists(strCRT_HostKeyFolder) Then 
						Call TrDebug ("CRT: FOUND DOESN'T EXIST: " & strCRT_HostKeyFolder,"", objDebug, MAX_LEN, 1, 1)								
						strCRT_HostKeyFolder = strCRT_ConfigFolder & "\Known Hosts"
						Call WriteStrToFile(SSH2File, "S:""Host Key Database Location""=" & strCRT_HostKeyFolder & "\", nSSH2, 1, 0)
						Call TrDebug ("CRT: NEW RECORD FOR HOST KEY FOLDER CREATED IN SSH2.ini: ","", objDebug, MAX_LEN, 1, 1)									
						RestartRequired = 3
						If Not objFSO.FolderExists(strCRT_HostKeyFolder) Then 
							Set objHostKeyFolder = objFSO.CreateFolder(strCRT_HostKeyFolder)
							Call TrDebug ("CRT: NEW HOSTKEY FOLDER: " & strCRT_HostKeyFolder,"CREATED", objDebug, MAX_LEN, 1, 1)														
							Set objHostKeyFolder = Nothing
						End If
					End If
				End If
			End If
		End If
		'------------------------------------------------------
		'  Copy Session Files to the SecureCrt Database
		'------------------------------------------------------
		If Not objFSO.FolderExists(strCRT_SessionFolder & "\" & strFolder) Then 
			objFSO.CreateFolder strCRT_SessionFolder & "\" & strFolder
		End If
		If objFSO.FileExists(strBinFolder & "\" & strSession & ".ini")	Then 
			objFSO.CopyFile strBinFolder & "\" & strSession & ".ini", strCRT_SessionFolder & "\" & strFolder & "\" & strSession & ".ini" , True
		End If
		If objFSO.FileExists(strBinFolder & "\" & strSessionSRX & ".ini")	Then 
			objFSO.CopyFile strBinFolder & "\" & strSessionSRX & ".ini", strCRT_SessionFolder & "\" & strFolder & "\" & strSessionSRX & ".ini" , True
		End If
		'-----------------------------------------------------
		'   CHECK IF PUBLICK KEYS ARE STORED for GATEWAY
		'-----------------------------------------------------
		'   GET IP ADDRESS FOR GATEWAY 1 SRX
		'-----------------------------------------------------
		Dim SRX_IP, SRX_Port, vSRX,vIncl,vExcl
		vIncl = Array("Hostname","[SSH2] Port")
		vExcl = Array("Null")
		nSSH2 = GetFileLineCountInclude(strCRT_SessionFolder & "\" & strFolder & "\" & strSessionSRX & ".ini", vSRX,vIncl, vExcl,0)
		If nSSH2 => 2 Then
			SRX_IP = Split(vSRX(0),"=")(1)
			SRX_Port = CLng("&h" & Split(vSRX(1),"=")(1))
			Call TrDebug ("CRT: SRX SESSION PARAMETERS: ",SRX_IP & ":" & SRX_Port, objDebug, MAX_LEN, 1, 1)
		Else
			Call TrDebug ("CRT: WRONG DATA IN SESSION FILE:" & strCRT_SessionFolder & "\" & strFolder & "\" & strSessionSRX & ".ini","ERROR", objDebug, MAX_LEN, 1, 1)
			crt.quit
			Exit Sub
		End If
		'-----------------------------------------------------
		'   GET IP ADDRESS FOR GATEWAY 2 SSG
		'-----------------------------------------------------
		Dim SSG_IP, SSG_Port, vSSG
		nSSH2 = GetFileLineCountInclude(strCRT_SessionFolder & "\" & strFolder & "\" & strSession & ".ini", vSSG,vIncl,vExcl,0)
		If nSSH2 => 2 Then
			SSG_IP = Split(vSSG(0),"=")(1)
			SSG_Port = CLng("&h" & Split(vSSG(1),"=")(1))
			Call TrDebug ("CRT: SSG SESSION PARAMETERS: ",SSG_IP & ":" & SSG_Port, objDebug, MAX_LEN, 1, 1)
		Else
			Call TrDebug ("CRT: WRONG DATA IN SESSION FILE:" & strCRT_SessionFolder & "\" & strFolder & "\" & strSessionSRX & ".ini","ERROR", objDebug, MAX_LEN, 1, 1)
			crt.quit
			Exit Sub
		End If
		'-----------------------------------------------------
		'   CHECK IF PUB KEY FOR SRX EXISTS
		'-----------------------------------------------------
		Dim SRXKeyFile, SSGKeyFile, SRXKeyFound, SSGKeyFound
		SRXKeyFound = False
		SSGKeyFound = False
		SRXKeyFile = SRX_IP & "[" & SRX_IP & "]" & SRX_Port & ".pub"
		SSGKeyFile = SSG_IP & "[" & SSG_IP & "]" & SSG_Port & ".pub"
		Call TrDebug ("CRT: LOOKING FOR: " & SRXKeyFile, "IN PROGRESS", objDebug, MAX_LEN, 1, 1)	
		Call TrDebug ("CRT: LOOKING FOR: " & SSGKeyFile, "IN PROGRESS", objDebug, MAX_LEN, 1, 1)		
		Set objHostKeyFolder = objFSO.GetFolder(strCRT_HostKeyFolder)
		Set colFiles = objHostKeyFolder.Files
		On Error Resume Next
			For Each objFile in colFiles
				If objFile.Name = SRXKeyFile Then SRXKeyFound = True : Call TrDebug ("CRT: " & SRXKeyFile, "FOUND", objDebug, MAX_LEN, 1, 1) : End If
				If objFile.Name = SSGKeyFile Then SSGKeyFound = True : Call TrDebug ("CRT: " & SSGKeyFile, "FOUND", objDebug, MAX_LEN, 1, 1) : End If	
			Next
		On Error Goto 0 
		Set objHostKeyFolder = Nothing
		Set colFiles = Nothing
		'---------------------------
		'   COPY HOSTKEY.PUB FILES
		'----------------------------
		If objFSO.FileExists(strBinFolder & "\" & SRXKeyFile)	Then 
			objFSO.CopyFile strBinFolder & "\" & SRXKeyFile, strCRT_HostKeyFolder & "\" & SRXKeyFile , True
			Call TrDebug ("CRT: FILE: " & SRXKeyFile, "COPIED", objDebug, MAX_LEN, 1, 1)	
		End If
		If objFSO.FileExists(strBinFolder & "\" & SSGKeyFile)	Then 
			objFSO.CopyFile strBinFolder & "\" & SSGKeyFile, strCRT_HostKeyFolder & "\" & SSGKeyFile , True
			Call TrDebug ("CRT: FILE: " & SSGKeyFile, "COPIED", objDebug, MAX_LEN, 1, 1)	
		End If
		If Not SRXKeyFound Then 
			Call WriteStrToFile(strCRT_HostKeyFolder & "\hostsmap.txt",SRX_IP & "[" & SRX_IP & "]" & SRX_Port & "=" & SRXKeyFile , 1, 2, 0)
			RestartRequired = 4
		End If 
		If Not SSGKeyFound Then 
			Call WriteStrToFile(strCRT_HostKeyFolder & "\hostsmap.txt",SSG_IP & "[" & SSG_IP & "]" & SSG_Port & "=" & SSGKeyFile , 1, 2, 0)			
			RestartRequired = 5			
		End If 
		Call WriteArrayToFile(strDirectoryLCL & "\Temp\"  & "crtupdate.ini",Array(GetScreenUserCrt()),1,1,0)								
		Call TrDebug ("CRT PARAMETERS FOR USER " & GetScreenUserCrt() & " UPDATED: " , "OK", objDebug, MAX_LEN, 3, 1)
		Exit Do
	Loop
	'-----------------------------------------------------
	'   CHECK IF RESTART IS REQUIRED
	'-----------------------------------------------------		
	If RestartRequired <> 0 Then 
	    Call TrDebug ("RESTART REQUIRED " , "RESTART CODE: " & RestartRequired, objDebug, MAX_LEN, 3, 1)
		'--------------------------------------------------------------------------------
		'           LOAD LIST OF SCRIPTS FROM inventory.ini
		'--------------------------------------------------------------------------------
			nInventory = GetFileLineCountByGroup(strDirectoryWork & "\" & strFileInventory, vInventory,"Client","","",0)
		'--------------------------------------------------------------------------------
		'          GET NAME OF THE TELNET SCRIPT
		'--------------------------------------------------------------------------------
		nIndex = GetObjectLineNumber(vInventory,nInventory,"TELNET", True)
		If nIndex = nInventory Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": ERROR:  CAN'T FIND NAME OF THE TELNET SCRIPT"  
				Exit Sub  
		End If
		setShedulerVBS = Split(vInventory(nIndex-1),",")(1) & "\" & Split(vInventory(nIndex-1),",")(0)
		strLine = " /ARG " & crt.Arguments(0) &_
			      " /ARG " & crt.Arguments(1) &_
			      " /ARG " & crt.Arguments(2) &_
			      " /ARG " & crt.Arguments(3) &_
			      " /SCRIPT " & strDirectoryWork & "\" & setShedulerVBS
		Call TrDebug ("RESTART SecureSRT With New Parameters ","", objDebug, MAX_LEN, 1, 1)			
		Call TrDebug ("RESTART: " & strCRTexe, "" , objDebug, MAX_LEN, Array(1,False), 1)					
		Call TrDebug ("RESTART: " & strLine, "" , objDebug, MAX_LEN, Array(1,False), 1)			
		objShell.run strCRTexe & strLine, nVisible, True
		set objShell = Nothing
		Set objFSO = Nothing
		crt.quit
        If IsObject(objDebug) Then objDebug.close : Set objDebug = Nothing End If		
	End If 
'------------------------------------------------------------------
'	START MAIN PROGRAM
'------------------------------------------------------------------
    Select Case crt.Arguments(0)
	    Case "ff", "ssg"
            vMod = Array("1","0","0","0","0")		    
		Case "combo"
		    vMod = Array("1","2","0","0","0")		    
		Case "srx"
		    vMod = Array("0","2","0","0","0")		    
		Case "closefw"
			vMod = Array("0","0","3","0","0")		    
		Case "checkfw"
			vMod = Array("0","0","0","4","0")	
        Case "srx_cmdlist"			
            vMod = Array("0","0","0","0","5")			
	End Select
	For each Mode in vMod
		Select Case Mode
		    Case "5" 
			        Call TrDebug ("CRT: LOAD CONFIGURATION", "", objDebug, MAX_LEN, 3, 1)
					If crt.Arguments.Count >= 4 Then strProcessName = crt.Arguments(3) Else strProcessName = "Null"
					nProgress = 65 + Int(My_Random(1,20))
                    Do
						Call TrDebug ("CONNECTING TO SRX", "", objDebug, MAX_LEN, 1, 1)
						If Not objFSO.FileExists(stdOutFile) Then 
							Call TrDebug ("ERROR CAN'T FIND TELNET COMMAND FILE", "", objDebug, MAX_LEN, 1, 1)
							If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "Abort013501")							
							Exit Do
						End If
						nSize = GetFileLineCountSelect(stdOutFile, vCommand,"","NULL","NULL",0)
						If nSize <= 0 Then 
							Call TrDebug ("CRT: NO COMMANDS NAMES FOUND IN " & stdOutFile, "ERROR", objDebug, MAX_LEN, 1, 1)
                            If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "100")							
							Exit Do
						End If 
						If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, nProgress)
						For nLine = 0 to UBound(vCommand) - 1
							Call TrDebug ("CRT: LOADING COMAND: " & vCommand(nLine), "", objDebug, MAX_LEN, 1, 1)
						Next
						'----------------------------------------------------------------
						'   CHECK COMMAND LIST FOR THE SAME COMANDS AND COPY TO NEW LIST
						'----------------------------------------------------------------
						Redim vCmdSrx(1) 
						strLine = "," : newIndex = 0
						For nInd = 0 to UBound(vCommand) - 1
						    If InStr(strLine,"," & vCommand(nInd) & ",") = 0 Then 
							    Redim Preserve vCmdSrx(newIndex + 1) 
								vCmdSrx(newIndex) = vCommand(nInd)
								newIndex = newIndex + 1
								strLine = strLine & vCommand(nInd) & ","
							End if
						Next
						If LoadSRXConfig(strFolder, strSessionSRX, strTab, vCommand,1) Then 
							If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "100")
						Else 
							If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "Abort013503")
						End If
						Exit Do
					Loop		
			Case "1"
			        Call TrDebug ("CONNECTING TO SSG", "", objDebug, MAX_LEN, 3, 1)
				'	stdOutFile = strDirectoryLCL & "\Temp\" & crt.Arguments(1)
				'	strTab = crt.Arguments(2)
					If objFSO.FileExists(stdOutFile) Then 
						Call GetFileLineCountSelect(stdOutFile, vCommand,"","NULL","NULL",0)
						Call TelnetGateway(strFolder, strSession, strTab, vCommand, 1)
						' objFSO.DeleteFile stdOutFile, True
					Else 
					    Call TrDebug ("ERROR CAN'T FIND TELNET COMMAND FILE", "", objDebug, MAX_LEN, 1, 1)
						crt.quit
						set objShell = Nothing
                		Set objFSO = Nothing
						If IsObject(objDebug) Then objDebug.close : Set objDebug = Nothing : End If
						Exit Sub
					End If
		    Case "2"
				'	strTab = crt.Arguments(2)
					Select Case Split(strTab,"-")(0)
					    Case "START"
						    strAction = "deactivate"
						Case "STOP" 
						    strAction = "activate"
					End Select
			        Call TrDebug ("NOW WILL " & strAction & " FILTER TERM FOR " & Split(strTab,"-")(1) , "", objDebug, MAX_LEN, 3, 1)
					If crt.Arguments.Count >= 4 Then strProcessName = crt.Arguments(3) Else strProcessName = "Null"
					nProgress = 65 + Int(My_Random(1,20))
					If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, nProgress)
					Call TrDebug ("CONNECTING TO SRX", "", objDebug, MAX_LEN, 3, 1)
				'	stdOutFile = strDirectoryLCL & "\Temp\" & crt.Arguments(1)
					
					If objFSO.FileExists(stdOutFile) Then 
						nSize = GetFileLineCountSelect(stdOutFile, vCommand,"","NULL","NULL",0)
						'--------------------------------------
						'   TRANSLATE CLI S2J
						'--------------------------------------
						Redim vCmdSrx(3 * nSize)
						For nLine = 0 to UBound(vCommand) - 1
						    vLine = Split(vCommand(nLine)," ")
						    If UBound(vLine) < 9 Then 
								Call TrDebug ("ERROR: CAN'T TRANSLATE CLI COMMAND", "", objDebug, MAX_LEN, 1, 1)
								crt.quit
								If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "Abort013001")
								set objShell = Nothing
                          		Set objFSO = Nothing
								If IsObject(objDebug) Then objDebug.close : Set objDebug = Nothing : End If
								Exit Sub
						    End If
						    StartTime = FormatDateTime(vLine(6), 4)
						    StopTime =   FormatDateTime(vLine(9), 4)
						    strDate = GetDateFormat(Date(),2)
						    Sched = Mid(vLine(2),2,Len(vLine(2)) - 2)
						    vCmdSrx(3 * nLine) = "delete schedulers scheduler " & Sched 
						    Call TrDebug ("CREATE COMMAND: " & vCmdSrx(3 * nLine), "", objDebug, MAX_LEN, 1, 1)
						    vCmdSrx(3 * nLine + 1) = "set schedulers scheduler " & Sched & " start-date " & strDate & "." & StartTime & " stop-date " & strDate & "." & StopTime
                            Call TrDebug ("CREATE COMMAND: " & vCmdSrx(3 * nLine + 1), "", objDebug, MAX_LEN, 1, 1)							
						    If InStr(Sched,"_TV") = 0  and InStr(Sched,"_Games") = 0 Then 
								vCmdSrx(3 * nLine + 2) = strAction & " firewall family inet filter FW_KIDSSHIELD term " & Sched
								Call TrDebug ("CREATE COMMAND: " & vCmdSrx(3 * nLine + 2), "", objDebug, MAX_LEN, 1, 1)
							Else 
							    vCmdSrx(3 * nLine + 2) = ""
								Call TrDebug ("CREATE COMMAND: Skip", "", objDebug, MAX_LEN, 1, 1)								
						    End If 
						Next
						If  LoadSRXConfig(strFolder, strSessionSRX, strTab, vCmdSrx, 1) Then 
							If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "100")
						Else 
							If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "Abort013003")
						End If
					Else 
					    Call TrDebug ("ERROR CAN'T FIND TELNET COMMAND FILE", "", objDebug, MAX_LEN, 1, 1)
						If strProcessName <> "Null" then Call CrtWriteProgressToFile(strProcessName, "Abort013002")
						crt.quit
						set objShell = Nothing
						Set objFSO = Nothing						
						If IsObject(objDebug) Then objDebug.close : Set objDebug = Nothing : End If						
						Exit Sub
					End If
		    Case "3"
				    Call TrDebug ("NOW WILL ACTIVATE FILTER TERM FOR " & Split(strTab,"-")(1), "", objDebug, MAX_LEN, 3, 1)
                    Do
						Select Case Split(strTab,"-")(0)
							Case "START"
								strAction = "deactivate"
							Case "STOP" 
								strAction = "activate"
						End Select
						Call TrDebug ("CONNECTING TO SRX", "", objDebug, MAX_LEN, 1, 1)
						If objFSO.FileExists(stdOutFile) Then 
							nSize = GetFileLineCountSelect(stdOutFile, vCommand,"","NULL","NULL",0)
							If nSize <= 0 Then 
							    Call TrDebug ("CRT: NO FILTER NAMES FOUND IN " & stdOutFile, "ERROR", objDebug, MAX_LEN, 1, 1)
								Exit Do
							End If 
							'--------------------------------------
							'   TRANSLATE CLI S2J
							'--------------------------------------
							' Redim vCmdSrx(nSize)
							For nLine = 0 to UBound(vCommand) - 1
								' Sched = vCommand(nLine)
								' vCmdSrx(nLine) = strAction & " firewall family inet filter FW_KIDSSHIELD term " & Sched 
								Call TrDebug ("CRT WILL CHECK SCHEDULERS: " & vCommand(nLine), "", objDebug, MAX_LEN, 1, 1)
							Next
							Call ActivateSRXFilters(strFolder, strSessionSRX, strTab, vCommand, "FW_KIDSSHIELD", 1)							
						Else 
							Call TrDebug ("ERROR CAN'T FIND TELNET COMMAND FILE", "", objDebug, MAX_LEN, 1, 1)
							Exit Do
						End If
						Exit Do
					Loop
			Case "4"
			        Call TrDebug ("NOW WILL CHECK FIREWALL TERMS CONSISTENCY", "", objDebug, MAX_LEN, 3, 1)
                    Do
						Select Case Split(strTab,"-")(0)
							Case "START"
								strAction = "deactivate"
							Case "STOP" 
								strAction = "activate"
						End Select
						Call TrDebug ("CONNECTING TO SRX", "", objDebug, MAX_LEN, 1, 1)
						If objFSO.FileExists(stdOutFile) Then 
							nSize = GetFileLineCountSelect(stdOutFile, vCommand,"","NULL","NULL",0)
							If nSize <= 0 Then 
							    Call TrDebug ("CRT: NO FILTER NAMES FOUND IN " & stdOutFile, "ERROR", objDebug, MAX_LEN, 1, 1)
								Exit Do
							End If 
							'--------------------------------------
							'   TRANSLATE CLI S2J
							'--------------------------------------
							' Redim vCmdSrx(nSize)
							For nLine = 0 to UBound(vCommand) - 1
								Call TrDebug ("CRT WILL CHECK SCHEDULERS: " & vCommand(nLine), "", objDebug, MAX_LEN, 1, 1)
							Next
							Call CheckSRXFilters(strFolder, strSessionSRX, strTab, vCommand, "FW_KIDSSHIELD", 1)	
						Else 
							Call TrDebug ("ERROR CAN'T FIND TELNET COMMAND FILE", "", objDebug, MAX_LEN, 1, 1)
							Exit Do
						End If
						Exit Do
					Loop		
		End Select		
	Next
	If IsObject(objDebug) Then objDebug.close : End If
	If objFSO.FileExists(stdOutFile) Then 
		objFSO.DeleteFile stdOutFile, True
	End If
	crt.quit
	Set objFSO = Nothing
    Set objShell = Nothing
End Sub
'----------------------------------------------------------------------------------
'    Function GetScreenUserCrt
'----------------------------------------------------------------------------------
Function GetScreenUserCrt()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserCrt = strScreenUser
End Function
'######################################################################
' Function TelnetGateway
'######################################################################
Function TelnetGateway(strFolder, strSession, strTab, ByRef vCommand, nDebug)		
	Dim  nResult, strStdOut, vStdOut
	Dim vCmd, nCmd
	Dim vCmdList
	Dim objTab
	Dim strCmd, strHost
	Const MAX_LEN = 130
	crt.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Start Telnet session to Home Gateway
    '--------------------------------------------------------------------------------
   nInd = 0
	'--------------------------------------------------------------------------------
	'	OPEN NEW SECURECRT TAB
	'--------------------------------------------------------------------------------
    On Error Resume Next
	Err.Clear
	Set objTab = crt.Session.ConnectInTab("/S " & strFolder & "\" & strSession)
	If Err.Number <> 0 Then 
		Call  TrDebug (strTab & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, 1, MAX_LEN, 1)
		TelnetGateway = False
		Exit Function
	End If
	On Error Goto 0
	'--------------------------------------------------------------------------------
	'	SEND <RETURN> KEY
	'--------------------------------------------------------------------------------
	objTab.Caption = strTab
	objTab.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Get actual host name of the box
    '--------------------------------------------------------------------------------
	objTab.Screen.Send chr(13)
	strLine = objTab.Screen.ReadString ("->")
    strHost = Split(strLine,chr(13))(1)
	Call TrDebug ("FOUND SSG GATEWAY NAME: " & strHost , "", objDebug, MAX_LEN, 1, nDebug)
	objTab.Screen.Send chr(13)
	nResult = objTab.Screen.WaitForString (strHost & "->",2)
    If nResult = 0  Then
	    Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM GATEWAY." , "", objDebug, MAX_LEN, 1, nDebug)
		TelnetGateway = False
		objTab.Session.Disconnect
    	Exit Function
    	Exit Function
	End If
	'--------------------------------------------------------------------------------
	'	SEND COMMANDS 
	'--------------------------------------------------------------------------------
	For i = 0 to UBound(vCommand) - 1
		strCmd = vCommand(i) 
		Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
		objTab.Screen.Send strCmd & chr(13)
		nResult = objTab.Screen.WaitForString (strHost & "->",2)
		nResult = objTab.Screen.WaitForString (strHost & "->",2)
		If nResult = 0  Then
			Call TrDebug ("WARNING: CAN'T GET RESPONSE FOR COMMAND " &  2 + i & " FROM GATEWAY.", "", objDebug, MAX_LEN, 1, nDebug)
			If IsObject(objDebug) Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & "  WARNING: CAN't GET RESPONSE FOR COMMAND " &  2 + i & " FROM GATEWAY. COMMAND MIGHT NOT APPLIED CORRECTLY" End If
		End If
	Next
	'--------------------------------------------------------------------------------
	'	SAVE CONFIGURATION
	'--------------------------------------------------------------------------------
	objTab.Screen.Send "save" & chr(13)
	nResult = objTab.Screen.WaitForString (strHost & "->",2)
    If nResult = 0  Then
		If IsObject(objDebug) Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & "  WARNING: CAN't GET RESPONSE FOR SAVE COMMAND. CONFIGURATION MIGHT NOT SAVED CORRECTLY" End If
	End If
    TelnetGateway = True
	objTab.Session.Disconnect
End Function
'-------------------------------------------------------------------------------------
' Function LoadSRXConfig
'-------------------------------------------------------------------------------------
Function LoadSRXConfig(strFolder, strSession, strTab, ByRef vCommand, nDebug)		
	Dim  nResult, strStdOut, vStdOut, nPrompt
	Dim vCmd, nCmd
	Dim vCmdList
	Dim objTab
	Dim strCmd, strHost
	Dim vWaitForCommit
	Dim Dstart, Tstart
	Const MAX_LEN = 130	
	vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
	crt.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Start Telnet session to Home Gateway
    '--------------------------------------------------------------------------------
    nInd = 0
	'--------------------------------------------------------------------------------
	'	OPEN NEW SECURECRT TAB
	'--------------------------------------------------------------------------------
    On Error Resume Next
	Err.Clear
	Set objTab = crt.Session.ConnectInTab("/S " & strFolder & "\" & strSession)
	If Err.Number <> 0 Then 
		Call  TrDebug (strTab & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, 1, MAX_LEN, 1)
		LoadSRXConfig = False
		Exit Function
	End If
	On Error Goto 0
	'--------------------------------------------------------------------------------
	'	SEND <RETURN> KEY
	'--------------------------------------------------------------------------------
	objTab.Caption = strTab
	objTab.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Get actual host name of the box
    '--------------------------------------------------------------------------------
	objTab.Screen.Send chr(13)
	strLine = objTab.Screen.ReadString (">")
    If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
	objTab.Screen.Send chr(13)
	nPrompt = objTab.Screen.WaitForString ("@" & strHost & ">",2)
    If nPrompt = 0  Then
		Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1)
		objTab.Session.Disconnect
        LoadSRXConfig = False
		Exit Function 
	End If	
	objTab.Screen.Send chr(13)
	objTab.Screen.WaitForString "@" & strHost & ">"
	objTab.Screen.Send "edit" & chr(13)
	objTab.Screen.WaitForString "@" & strHost & "#"
	'--------------------------------------------------------------------------------
	'	SEND COMMANDS 
	'--------------------------------------------------------------------------------
	For i = 0 to UBound(vCommand) - 1
		strCmd = vCommand(i) 
		Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
		objTab.Screen.Send strCmd & chr(13)
        objTab.Screen.WaitForString "@" & strHost & "#"
	Next
	'--------------------------------------------------------------------------------
	'	COMMIT CONFIGURATION
	'--------------------------------------------------------------------------------
	Dstart = Date()
	Tstart = Time()
	Call  TrDebug ("COMMIT " & strHost, "......IN PROGRESS", objDebug, MAX_LEN, 1, 1)   
	objTab.Screen.Send "commit" & chr(13)
	nResult = objTab.Screen.WaitForStrings (vWaitForCommit, 30)
    Select Case nResult
        Case 0
			Call  TrDebug ("COMMIT " & strHost, "TIME OUT", objDebug, MAX_LEN, 1, 1) 
            LoadSRXConfig = False
	        objTab.Session.Disconnect			
			Exit Function 
        Case 1 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 1", objDebug, MAX_LEN, 1, 1)
			For i = 0 to 3
			    nPrompt = objTab.Screen.WaitForString ("@" & strHost & "#",2)
				Select Case nPrompt
					Case 0
						Call TrDebug ("ERROR: CAN'T GET PROMPT(" & i & ") FROM NODE", "", objDebug, MAX_LEN, 1, 1)
						If i = 3 Then 
							objTab.Session.Disconnect
							LoadSRXConfig = False
							Call TrDebug ("LOST CONNECTION. ", "ERROR", objDebug, MAX_LEN, 1, 1)
							Exit Function
						End If 
						objTab.Screen.Send chr(13)
					Case Else 
						objTab.Screen.Send "rollback" & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
						Call TrDebug ("CONFIGURATION ROLLBACK", "OK", objDebug, MAX_LEN, 1, 1)
						If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
						objTab.Session.Disconnect			
						LoadSRXConfig = False
						Exit Function			
				End Select	
            Next			
        Case 2 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 2", objDebug, MAX_LEN, 1, 1)
			For i = 0 to 3
			    nPrompt = objTab.Screen.WaitForString ("@" & strHost & "#",2)
				Select Case nPrompt
					Case 0
						Call TrDebug ("ERROR: CAN'T GET PROMPT(" & i & ") FROM NODE", "", objDebug, MAX_LEN, 1, 1)
						If i = 3 Then 
							objTab.Session.Disconnect
							LoadSRXConfig = False
							Call TrDebug ("LOST CONNECTION. ", "ERROR", objDebug, MAX_LEN, 1, 1)
							Exit Function
						End If 
						objTab.Screen.Send chr(13)
					Case Else 
						objTab.Screen.Send "rollback" & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
						Call TrDebug ("CONFIGURATION ROLLBACK", "OK", objDebug, MAX_LEN, 1, 1)
						If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
						objTab.Session.Disconnect			
						LoadSRXConfig = False
						Exit Function			
				End Select	
            Next			
		Case Else
			Call  TrDebug ("COMMIT " & strHost, "OK", objDebug, MAX_LEN, 1, 1)   
    End Select		
	Tcompiler = DateDiff("s",Dstart & " " & Tstart,Date() & " " & Time()) 
	Call TrDebug("Commit time: " & Tcompiler & " sec" ,"", objDebug, MAX_LEN, 1, 1)
    LoadSRXConfig = True
	objTab.Screen.WaitForString "@" & strHost & "#"
	If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
	objTab.Session.Disconnect
End Function
'------------------------------------------------------------------------------------------
'    Function ActivateSRXFilters(strFolder, strSession, Byref strHost, strTab, ByRef vCommand, strFilter, nDebug)		
'-------------------------------------------------------------------------------------------
Function ActivateSRXFilters(strFolder, strSession, strTab, ByRef vCommand, strFilter, nDebug)		
	Dim  nResult, strStdOut, vStdOut, nPrompt
	Dim vCmd, nCmd
	Dim vCmdList
	Dim objTab
	Dim strCmd, strHost
	Dim vWaitForCommit, NothingToActivate, vTerm, nTerm
	Const MAX_LEN = 130	
	Redim vTerm(0)
	vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
	crt.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Start Telnet session to Home Gateway
    '--------------------------------------------------------------------------------
   nInd = 0
	'--------------------------------------------------------------------------------
	'	OPEN NEW SECURECRT TAB
	'--------------------------------------------------------------------------------
    On Error Resume Next
	Err.Clear
	Set objTab = crt.Session.ConnectInTab("/S " & strFolder & "\" & strSession)
	If Err.Number <> 0 Then 
		Call  TrDebug (strTab & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, 1, MAX_LEN, 1)
		ActivateSRXFilters = False
		Exit Function
	End If
	On Error Goto 0
	'--------------------------------------------------------------------------------
	'	SEND <RETURN> KEY
	'--------------------------------------------------------------------------------
	objTab.Caption = strTab
	objTab.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Get actual host name of the box
    '--------------------------------------------------------------------------------
	objTab.Screen.Send chr(13)
	strLine = objTab.Screen.ReadString (">")
    If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
	objTab.Screen.Send chr(13)
	nResult = objTab.Screen.WaitForString ("@" & strHost & ">",2)
    If nResult = 0  Then
		Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1)
		objTab.Session.Disconnect
        ActivateSRXFilters = False
		Exit Function 
	End If	
	objTab.Screen.Send chr(13)
	objTab.Screen.WaitForString "@" & strHost & ">"
	objTab.Screen.Send "edit" & chr(13)
	objTab.Screen.WaitForString "@" & strHost & "#"
	'--------------------------------------------------------------------------------
	'	CHECK CURRENT FILTER STATUS
	'--------------------------------------------------------------------------------
	Call TrDebug ("CHECK FILTER TERM STATUS: " & "show firewall family inet filter " & strFilter , "", objDebug, MAX_LEN, 1, 1)
    objTab.Screen.Send "show firewall family inet filter " & strFilter & " |display set |match deactivate" & chr(13)
	strLine = objTab.Screen.ReadString ("@" & strHost & "#")
	NeedToActivate = False
	nTerm = 0
	For i = 0 to UBound(vCommand) - 1
	    If InStr(strLine,vCommand(i)) <> 0 Then
		   NeedToActivate = True
		   Redim Preserve vTerm(nTerm + 1)
		   vTerm(nTerm) = vCommand(i)
		   Call TrDebug ("ActivateFilter: FOUND INACTIVE TERM:[" & vTerm(nTerm) & "]", "", objDebug, MAX_LEN, 1, 1)
		   nTerm = nTerm + 1
		End If
    Next
	If Not NeedToActivate Then
		Call TrDebug("All Filters are alredy active " ,"EXIT", objDebug, MAX_LEN, 1, 1)
		ActivateSRXFilters = True	
		If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
		objTab.Session.Disconnect
		Exit Function
	End If 
	'--------------------------------------------------------------------------------
	'	SEND COMMANDS 
	'--------------------------------------------------------------------------------
	For i = 0 to UBound(vTerm) - 1
		strCmd = "activate firewall family inet filter " & strFilter & " term " & vTerm(i)
		Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
		objTab.Screen.Send strCmd & chr(13)
        objTab.Screen.WaitForString "@" & strHost & "#"
	Next
	'--------------------------------------------------------------------------------
	'	COMMIT CONFIGURATION
	'--------------------------------------------------------------------------------
	Dstart = Date()
	Tstart = Time()
	Call  TrDebug ("COMMIT " & strHost, "......IN PROGRESS", objDebug, MAX_LEN, 1, 1)   
	objTab.Screen.Send "commit" & chr(13)
	nResult = objTab.Screen.WaitForStrings (vWaitForCommit, 40)
    Select Case nResult
        Case 0
			Call  TrDebug ("COMMIT " & strHost, "TIME OUT", objDebug, MAX_LEN, 1, 1) 
            ActivateSRXFilters = False
	        objTab.Session.Disconnect			
			Exit Function 
        Case 1 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 1", objDebug, MAX_LEN, 1, 1)
			For i = 0 to 3
			    nPrompt = objTab.Screen.WaitForString ("@" & strHost & "#",2)
				Select Case nPrompt
					Case 0
						Call TrDebug ("ERROR: CAN'T GET PROMPT(" & i & ") FROM NODE", "", objDebug, MAX_LEN, 1, 1)
						If i = 3 Then 
							objTab.Session.Disconnect
							ActivateSRXFilters = False
							Call TrDebug ("LOST CONNECTION. ", "ERROR", objDebug, MAX_LEN, 1, 1)
							Exit Function
						End If 
						objTab.Screen.Send chr(13)
					Case Else 
						objTab.Screen.Send "rollback" & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
						Call TrDebug ("CONFIGURATION ROLLBACK", "OK", objDebug, MAX_LEN, 1, 1)
						If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
						objTab.Session.Disconnect			
						ActivateSRXFilters = False
						Exit Function			
				End Select	
            Next			
        Case 2 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 2", objDebug, MAX_LEN, 1, 1)
			For i = 0 to 3
			    nPrompt = objTab.Screen.WaitForString ("@" & strHost & "#",2)
				Select Case nPrompt
					Case 0
						Call TrDebug ("ERROR: CAN'T GET PROMPT(" & i & ") FROM NODE", "", objDebug, MAX_LEN, 1, 1)
						If i = 3 Then 
							objTab.Session.Disconnect
							ActivateSRXFilters = False
							Call TrDebug ("LOST CONNECTION. ", "ERROR", objDebug, MAX_LEN, 1, 1)
							Exit Function
						End If 
						objTab.Screen.Send chr(13)
					Case Else 
						objTab.Screen.Send "rollback" & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
						Call TrDebug ("CONFIGURATION ROLLBACK", "OK", objDebug, MAX_LEN, 1, 1)
						If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
						objTab.Session.Disconnect			
						ActivateSRXFilters = False
						Exit Function			
				End Select	
            Next			
		Case Else
			Call  TrDebug ("COMMIT " & strHost, "OK", objDebug, MAX_LEN, 1, 1)   
    End Select		
	Tcompiler = DateDiff("s",Dstart & " " & Tstart,Date() & " " & Time()) 
	Call TrDebug("Commit time: " & Tcompiler & " sec" ,"", objDebug, MAX_LEN, 1, 1)
    ActivateSRXFilters = True
	objTab.Screen.WaitForString "@" & strHost & "#"
    If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
	objTab.Session.Disconnect
End Function
'------------------------------------------------------------------------------------------
'    Function CheckSRXFilters(strFolder, strSession, Byref strHost, strTab, ByRef vCommand, strFilter, nDebug)		
'-------------------------------------------------------------------------------------------
Function CheckSRXFilters(strFolder, strSession, strTab, ByRef vCommand, strFilter, nDebug)		
	Dim  nResult, strStdOut, vStdOut
	Dim vCmd, nCmd
	Dim vCmdList
	Dim objTab
	Dim strCmd, strHost
	Dim vWaitForCommit, vTerm, vLine
	Dim CommitRequired, ShouldBeInactive
	Const MAX_LEN = 130	
	Redim vTerm(0)
	vWaitForCommit = Array("error: configuration check-out failed","error: commit failed","commit complete")
	crt.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Start Telnet session to Home Gateway
    '--------------------------------------------------------------------------------
   nInd = 0
	'--------------------------------------------------------------------------------
	'	OPEN NEW SECURECRT TAB
	'--------------------------------------------------------------------------------
    On Error Resume Next
	Err.Clear
	Set objTab = crt.Session.ConnectInTab("/S " & strFolder & "\" & strSession)
	If Err.Number <> 0 Then 
		Call  TrDebug (strTab & "ERROR:", Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description , objDebug, 1, MAX_LEN, 1)
		CheckSRXFilters = False
		Exit Function
	End If
	On Error Goto 0
	'--------------------------------------------------------------------------------
	'	SEND <RETURN> KEY
	'--------------------------------------------------------------------------------
	objTab.Caption = strTab
	objTab.Screen.Synchronous = True
	'--------------------------------------------------------------------------------
    '  Get actual host name of the box
    '--------------------------------------------------------------------------------
	objTab.Screen.Send chr(13)
	strLine = objTab.Screen.ReadString (">")
    If InStr(strLine,"@") Then strHost = Split(strLine,"@")(1)
	objTab.Screen.Send chr(13)
	nResult = objTab.Screen.WaitForString ("@" & strHost & ">",2)
    If nResult = 0  Then
		Call TrDebug ("ERROR: CAN'T GET RESPONSE FROM NODE", "", objDebug, MAX_LEN, 1, 1)
		objTab.Session.Disconnect
        CheckSRXFilters = False
		Exit Function 
	End If	
	objTab.Screen.Send chr(13)
	objTab.Screen.WaitForString "@" & strHost & ">"
	objTab.Screen.Send "edit" & chr(13)
	objTab.Screen.WaitForString "@" & strHost & "#"
	'--------------------------------------------------------------------------------
	'	CHECK INACTIVE FILTERS STATUS
	'--------------------------------------------------------------------------------
	objTab.Screen.Send "show firewall family inet filter " & strFilter & " |match ""inactive: term""" & chr(13)
	strLine = objTab.Screen.ReadString ("@" & strHost & "#")
	vLine = Split(strLine,chr(13))
	nTerm = 0 
	For i = 0 to UBound(vLine)
		If InStr(vLine(i),"inactive: term ") > 0 Then 
		   Redim Preserve vTerm(nTerm + 1)
		   vTerm(nTerm) = Split(vLine(i)," ")(2)
		   Call TrDebug ("CheckFilter: FOUND INACTIVE TERM:[" & vTerm(nTerm) & "]", "", objDebug, MAX_LEN, 1, 1)
		   nTerm = nTerm + 1
		End If
	Next
	'--------------------------------------------------------------------------------
	'	SEND COMMANDS 
	'--------------------------------------------------------------------------------
    For nTerm = 0 to UBound(vTerm) - 1
	    ShouldBeInactive = False
		For i = 0 to UBound(vCommand) - 1
		    If vTerm(nTerm) = vCommand(i) Then ShouldBeInactive = True
		Next
		If Not ShouldBeInactive Then 
   			strCmd = "activate firewall family inet filter " & strFilter & " term " & vTerm(nTerm)
			Call TrDebug ("TERM: " &  vTerm(nTerm)," INCONSISTENT", objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
			objTab.Screen.Send strCmd & chr(13)
			objTab.Screen.WaitForString "@" & strHost & "#"
			CommitRequired = True
		Else 
		    Call TrDebug ("TERM: " &  vTerm(nTerm)," CONSISTENT", objDebug, MAX_LEN, 1, nDebug)
		End If
	Next
	'--------------------------------------------------------------------------------
	'	CHECK ACTIVE FILTERS STATUS
	'--------------------------------------------------------------------------------
	Redim vTerm(1)
	objTab.Screen.Send "show firewall family inet filter " & strFilter & " |match term |except inactive:" & chr(13)
	strLine = objTab.Screen.ReadString ("@" & strHost & "#")
	vLine = Split(strLine,chr(13))
	nTerm = 0 
	For i = 0 to UBound(vLine)
		If InStr(vLine(i),"term ") and InStr(vLine(i),"inactive:") = 0 Then 
		   Redim Preserve vTerm(nTerm + 1)
		   vTerm(nTerm) = Split(vLine(i)," ")(1)
		   Call TrDebug ("CheckFilter: FOUND ACTIVE TERM:[" & vTerm(nTerm) & "]", "", objDebug, MAX_LEN, 1, 1)
		   nTerm = nTerm + 1
		End If
	Next
	'--------------------------------------------------------------------------------
	'	SEND COMMANDS 
	'--------------------------------------------------------------------------------
    For nTerm = 0 to UBound(vTerm) - 1
	    ShouldBeInactive = False
		For i = 0 to UBound(vCommand) - 1
		    If vTerm(nTerm) = vCommand(i) Then ShouldBeInactive = True
		Next
		If ShouldBeInactive Then 
   			strCmd = "deactivate firewall family inet filter " & strFilter & " term " & vTerm(nTerm)
			Call TrDebug ("TERM: " &  vTerm(nTerm)," INCONSISTENT", objDebug, MAX_LEN, 1, nDebug)
			Call TrDebug ("SEND COMMAND: " & strCmd , "", objDebug, MAX_LEN, 1, nDebug)
			objTab.Screen.Send strCmd & chr(13)
			objTab.Screen.WaitForString "@" & strHost & "#"
			CommitRequired = True
		Else 
		    Call TrDebug ("TERM: " &  vTerm(nTerm)," CONSISTENT", objDebug, MAX_LEN, 1, nDebug)
		End If
	Next
	'--------------------------------------------------------------------------------
	'	COMMIT CONFIGURATION
	'--------------------------------------------------------------------------------
	If CommitRequired Then
		Dstart = Date()
		Tstart = Time()
		Call  TrDebug ("COMMIT " & strHost, "......IN PROGRESS", objDebug, MAX_LEN, 1, 1)   
		objTab.Screen.Send "commit" & chr(13)
		nResult = objTab.Screen.WaitForStrings (vWaitForCommit, 30)
		Select Case nResult
			Case 0
				Call  TrDebug ("COMMIT " & strHost, "TIME OUT", objDebug, MAX_LEN, 1, 1) 
				CheckSRXFilters = False
				objTab.Session.Disconnect			
				Exit Function 
        Case 1 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 1", objDebug, MAX_LEN, 1, 1)
			For i = 0 to 3
			    nPrompt = objTab.Screen.WaitForString ("@" & strHost & "#",2)
				Select Case nPrompt
					Case 0
						Call TrDebug ("ERROR: CAN'T GET PROMPT(" & i & ") FROM NODE", "", objDebug, MAX_LEN, 1, 1)
						If i = 3 Then 
							objTab.Session.Disconnect
							CheckSRXFilters = False
							Call TrDebug ("LOST CONNECTION. ", "ERROR", objDebug, MAX_LEN, 1, 1)
							Exit Function
						End If 
						objTab.Screen.Send chr(13)
					Case Else 
						objTab.Screen.Send "rollback" & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
						Call TrDebug ("CONFIGURATION ROLLBACK", "OK", objDebug, MAX_LEN, 1, 1)
						If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
						objTab.Session.Disconnect			
						CheckSRXFilters = False
						Exit Function			
				End Select	
            Next			
        Case 2 
			Call  TrDebug ("COMMIT " & strHost, "ERROR 2", objDebug, MAX_LEN, 1, 1)
			For i = 0 to 3
			    nPrompt = objTab.Screen.WaitForString ("@" & strHost & "#",2)
				Select Case nPrompt
					Case 0
						Call TrDebug ("ERROR: CAN'T GET PROMPT(" & i & ") FROM NODE", "", objDebug, MAX_LEN, 1, 1)
						If i = 3 Then 
							objTab.Session.Disconnect
							CheckSRXFilters = False
							Call TrDebug ("LOST CONNECTION. ", "ERROR", objDebug, MAX_LEN, 1, 1)
							Exit Function
						End If 
						objTab.Screen.Send chr(13)
					Case Else 
						objTab.Screen.Send "rollback" & chr(13)
						objTab.Screen.WaitForString "@" & strHost & "#"
						Call TrDebug ("CONFIGURATION ROLLBACK", "OK", objDebug, MAX_LEN, 1, 1)
						If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
						objTab.Session.Disconnect			
						CheckSRXFilters = False
						Exit Function			
				End Select	
            Next			
			Case Else
				Call  TrDebug ("COMMIT " & strHost, "OK", objDebug, MAX_LEN, 1, 1)   
		End Select		
		Tcompiler = DateDiff("s",Dstart & " " & Tstart,Date() & " " & Time()) 
		Call TrDebug("Commit time: " & Tcompiler & " sec" ,"", objDebug, MAX_LEN, 1, 1)
	    objTab.Screen.WaitForString "@" & strHost & "#"
    End If 
	CheckSRXFilters = True
	If Not LeaveJunos(objTab) Then Call TrDebug ("EXIT TIMEOUT", "ERROR", objDebug, MAX_LEN, 1, 1)
	objTab.Session.Disconnect
End Function
'----------------------------------------------------------------
' Function CrtWriteProgressToFile 
'----------------------------------------------------------------
 Function CrtWriteProgressToFile(strProcessName, pProgress)
	Dim objProgressFile, f_objFSO, g_objShell
	Set g_objShell = CreateObject("WScript.Shell")	
	strWork = g_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	Set objProgressFile = f_objFSO.OpenTextFile(strWork & "\" & strProcessName & ".dat",ForAppending,True)
		objProgressFile.WriteLine pProgress
    objProgressFile.close
	Set f_objFSO = Nothing
	Set g_objShell = Nothing
End Function
'---------------------------------------------------------------------------------
' Function LeaveJunos(objTab) 
'---------------------------------------------------------------------------------
Function LeaveJunos(objTab)
Dim vWaitForExit, nResult
    LeaveJunos = False
    vWaitForExit = Array(">","(yes)")
	objTab.Screen.Send "exit" & chr(13)
	nResult = objTab.Screen.WaitForStrings (vWaitForExit, 5)
	Select Case nResult
	    Case 0
		     LeaveJunos = False
		Case 1
		     LeaveJunos = True
		Case 2
		    objTab.Screen.Send "yes" & chr(13)
		    objTab.Screen.WaitForString ">"
			LeaveJunos = True
	End select
End Function
'---------------------------------------------------------------------------------------
' 	nMode = 2  Then Insert Above
'   nMode = 3  Then Insert Below
' 	nMode = 1  Then Change
'   nMode = 4  Append Line to File
'	Inserts or change Line in Text File at String Number "LineNumber" (count form 1)
'   Function WriteStrToFile(strDirectoryTmp & "\" & strFileLocalSessionTmp, nTime, LineNumber, CHANGE)
'---------------------------------------------------------------------------------------
Function WriteStrToFile(strFile, strNewLine, LineNumber, nMode, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Dim objFSO, objFile
	Const FOR_WRITING = 1
	WriteStrToFile = False
	If LineNumber > 10000 Then objDebug.WriteLine "WriteStrToFile: ERROR: CAN'T OPERATE FILES WITH MORE THEN 1000 STRINGS" End If  
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objFile = objFSO.OpenTextFile(strFile,2,True)
		objFile.WriteLine "Empty"
		If Err.Number = 0 Then 
			objFile.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLIne is number of lines counted like 1,2,...,n
	If nMode = 2 and LineNumber > nFileLine Then nMode = 12 End If
	If nMode = 3 and LineNumber > nFileLine Then nMode = 12 End If	
	If nMode = 1 and LineNumber > nFileLine Then nMode = 12 End If
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": WriteStrToFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	Select Case nMode
			Case 1 																		' - CHANGE REQUESTED LINENUMBER
					vFileLine(LineNumber - 1) = strNewLine
					If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 2
					Redim vvFileLine(nFileLine + 1)
					For i = 0 to LineNumber - 2
						vvFileLine(i) = vFileLine(i)
					Next
					vvFileLine(LineNumber - 1) = strNewLine
					For i = LineNumber to nFileLine
						vvFileLine(i) = vFileLine(i-1)
					Next
					nFileLine = nFileLine + 1
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 3 ' - Insert After
					Redim vvFileLine(nFileLine + 1)
					For i = 0 to LineNumber - 1
						vvFileLine(i) = vFileLine(i)
					Next
					vvFileLine(LineNumber) = strNewLine
					For i = LineNumber + 1 to nFileLine
						vvFileLine(i) = vFileLine(i-1)
					Next
					nFileLine = nFileLine + 1
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
			Case 4 ' - Append string to file
			        Set objFile = objFSO.OpenTextFile(strFile,8,True)
					objFile.WriteLine strNewLine
					objFile.Close
			Case 12
					Redim vvFileLine(LineNumber)
					For i = 0 to nFileLine - 1
						vvFileLine(i) = vFileLine(i)
					Next
					For i = nFileLine to LineNumber - 2
						vvFileLine(i) = " "
					Next
					vvFileLine(LineNumber - 1) = strNewLine
					nFileLine = LineNumber
					If WriteArrayToFile(strFile,vvFileLine,nFileLine,FOR_WRITING,nDebug) Then WriteStrToFile = True End If
	End Select
	Set objFSO = Nothing
	Set objFile = Nothing
End Function
 '#######################################################################
 ' Creates File if it doesn't exists
 ' nMode = 2  Then Append
 ' nMode = 1  Then Rewire all File content
 ' Function WriteArrayToFile - Returns number of lines int the text file
 '#######################################################################
 Function WriteArrayToFile(strFile,vFileLine, nFileLine,nMode,nDebug)
    Dim i, nCount
	Dim strLine
	Dim objDataFileName, objFSO
	
	WriteArrayToFile = False
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(strFile) Then 	
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.CreateTextFile(strFile)
		If Err.Number = 0 Then 
			objDataFileName.close
			On Error Goto 0
		Else
			Set objFSO = Nothing
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T CREATE FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			On Error Goto 0
			Exit Function
		End If
	End If
	
	Select Case nMode
		Case 1 
			Set objDataFileName = objFSO.OpenTextFile(strFile,2,True)
		Case 2 	
			Set objDataFileName = objFSO.OpenTextFile(strFile,8,True)
	End Select 
	i = 0
	On Error Resume Next
	Err.Clear
	Do While i < nFileLine
		objDataFileName.WriteLine vFileLine(i)
		If Err.Number <> 0 Then 
			If IsObject(objDebug) Then 
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile: ERROR: CAN'T WRITE TO FILE " & strFile
				objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WriteArrayToFile:  Error: " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description
			End If
			WriteArrayToFile = False
			Exit Do 			
		End If
		i = i + 1
	Loop
	On Error Goto 0
	If i = nFileLine Then WriteArrayToFile = True End If
	objDataFileName.close
	Set objFSO = Nothing
End Function
 '#######################################################################
 ' Function GetFileLineCountByGroup - Returns number of lines int the text file
'#######################################################################
 Function GetFileLineCountByGroup(strFileName, ByRef vFileLines, strGroup1, strGroup2, strGroup3, nDebug_)
    Dim nIndex
	Dim strLine 
	Dim nGroupSelector
	Dim vDelim
	vDelim = Array("=",",",":")
	GetFileLineCountByGroup = 0
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	nIndex = 0
	Redim vFileLines(nIndex)
	If nDebug_ = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: -------------- NOW READ TO ARRAY -----------" End If  
	nGroupSelector = 0
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
    Do While objDataFileName.AtEndOfStream <> True
		' strLine = All_Trim(objDataFileName.ReadLine)
		strLine = RTrim(LTrim(objDataFileName.ReadLine))
		Select Case Left(strLine,1)
			Case "#"
			Case "$"
			Case ""
			Case "["
				If strGroup1 = "All" Then 
					nGroupSelector = 1 
				Else 
					Select Case strLine
						Case "[" & strGroup1 & "]"
							nGroupSelector = 1
						Case "[" & strGroup2 & "]"
							nGroupSelector = 1
						Case "[" & strGroup3 & "]"
							nGroupSelector = 1
						Case Else
							nGroupSelector = 0
					End Select
				End If
			Case Else	
				If nGroupSelector = 1 Then
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = NormalizeStr(strLine, vDelim)
					If nDebug_ = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": GetFileLineCountByGroup: vFileLines(" & nIndex & "): "  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
				End If
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountByGroup = nIndex
End Function
'#######################################################################
' Function GetFileLineCountSelect - Returns number of lines int the text file
'#######################################################################
 Function GetFileLineCountSelect(strFileName, ByRef vFileLines,strChar1, strChar2, strChar3, nDebug)
    Dim nIndex
	Dim strLine
	Dim objDataFileName
    strFileWeekStream = ""	
	If objFSO.FileExists(strFileName) Then 
		On Error Resume Next
		Err.Clear
		Set objDataFileName = objFSO.OpenTextFile(strFileName)
		If Err.Number <> 0 Then 
			Call TrDebug("GetFileLineCountSelect: ERROR: CAN'T OPEN FILE:", strFileName, objDebug, MAX_LEN, 0, 1)
			On Error Goto 0
			Redim vFileLines(0)
			GetFileLineCountSelect = 0
			Exit Function
		End If
	Else
	    Call TrDebug("GetFileLineCountSelect: ERROR: CAN'T FIND FILE:", strFileName, objDebug, MAX_LEN, 0, 1)
		Redim vFileLines(0)
		GetFileLineCountSelect = 0
		Exit Function
	End If 
    Redim vFileLines(0)
	Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	If nDebug = 1 Then objDebug.WriteLine "           NOW TRYING TO RIGHT INTO AN ARRAY        "
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
		strLine = objDataFileName.ReadLine
		Select Case Left(strLine,1)
			Case strChar1
			Case strChar2
			Case strChar3
			Case Else
					Redim Preserve vFileLines(nIndex + 1)
					vFileLines(nIndex) = strLine
					If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
					nIndex = nIndex + 1
		End Select
	Loop
	objDataFileName.Close
    GetFileLineCountSelect = nIndex
End Function
'-----------------------------------------------------------------
'     Function Normalize(strLine) - Removes all spaces around delimiters: arg1...arg3
'-----------------------------------------------------------------
Function NormalizeStr(strLine, vDelim)
Dim strNew
	strLine = LTrim(RTrim(strLine))
	strNew = ""
	For nInd = 0 to UBound(vDelim) - 1
		i = 0 
		Do While i <= UBound(Split(strLine,vDelim(nInd)))
				If UBound(Split(strLine,vDelim(nInd))) = 0 Then Exit Do End If
				If i < UBound(Split(strLine,vDelim(nInd))) Then strNew = strNew & LTrim(RTrim(Split(strLine,vDelim(nInd))(i))) & vDelim(nInd) End If
				If i = UBound(Split(strLine,vDelim(nInd))) Then strNew = strNew & LTrim(RTrim(Split(strLine,vDelim(nInd))(i))) End If
				i = i + 1
		Loop
		If i > 0 Then strLine = strNew End If
		strNew = ""
	Next
	NormalizeStr = strLine
End Function 
'-----------------------------------------------------------------
'     Function All_Trim(strLine) - Removes all speces form the string
'-----------------------------------------------------------------
Function All_Trim(strLine)
Dim nChar, strChar, i, strResult
	strResult = ""
	nChar = Len(strLine)
	For i = 1 to nChar
		strChar = Mid(strLine,i,1)
		If strChar <> " " Then strResult = strResult & strChar End If
	Next
		All_Trim = strResult
End Function
'------------------------------------------------------------------------------------------------------------------
' Function returns the number of the line from 1 to N which contains string strObject. Returns 0 if nothing found
'------------------------------------------------------------------------------------------------------------------
Function GetObjectLineNumber( byRef vArray, nArrayLen, strObjectName, bCaseSensitive)
 Dim nInd, strLine, strPattern
	nInd = 0
	GetObjectLineNumber = 0
	Do While nInd < nArrayLen
		Select Case bCaseSensitive
		    Case True
		        strLine =  vArray(nInd)
                strPattern = strObjectName				
		    Case False
   		        strLine =  LCase(vArray(nInd))
                strPattern = LCase(strObjectName)
		End Select
		If InStr(strLine, strPattern) <> 0	Then 
			GetObjectLineNumber = nInd + 1
			Exit Do
		End If
		nInd = nInd + 1
    Loop
End Function
'##############################################################################################################
'--------------------------------------------------------------------------------
' Function GetFileLineCountInclude - Returns number of lines int the text file
'--------------------------------------------------------------------------------
Function GetFileLineCountInclude(strFileName, ByRef vFileLines,vIncl, vExcl,nDebug)
    Dim nIndex
	Dim strLine, nCount, nSize
	Dim objDataFileName, nResult, nSymbol
	
    strFileWeekStream = ""	
	Set objDataFileName = objFSO.OpenTextFile(strFileName)
	Redim vFileLines(0)
	Set objDataFileName = objFSO.OpenTextFile(strFileName)	
	nIndex = 0
    Do While objDataFileName.AtEndOfStream <> True
	    ' Check if string contains symbol or substric to exclude from coping to array
		nResult = True
		strLine = LTrim(objDataFileName.ReadLine)
        ' Additional check of the fully commented lines
        if InStr(strLine,"'") Then 
		    nSymbol = 1
			Do 
			    strChar = Mid(strLine,nSymbol,1)
			    Select Case strChar
				    Case "'"
						nResult = False
						Exit Do
					Case chr(9)
					    nSymbol = nSymbol + 1
					Case Else 
						nResult = True
						Exit Do
			    End Select
           Loop			
        End If 		
        ' Check if string contains symbol or substring which must be excluded 
	    If nResult = True and  vExcl(0) <> "Null" Then 
			nExcl = 0
		    Do while nExcl <= UBound(vExcl)
		        If InStr(strLine, vExcl(nExcl)) <> 0 Then 
				    nResult = False
					Exit Do
				End If 
				nExcl = nExcl + 1
		    Loop
		End If
        ' Check if string contains symbol or substring which must be included 
    	If nResult = True and vIncl(0) <> "All" Then 
			nResult = false
			nIncl = 0
			Do while nIncl <= UBound(vIncl)
				If InStr(strLine, vIncl(nIncl)) <> 0 Then 
					
					nResult = True
					Exit Do
				End If 
				nIncl = nIncl + 1
			Loop
		End If 
        If nResult Then 
			Redim Preserve vFileLines(nIndex + 1)
			vFileLines(nIndex) = strLine
			If nDebug = 1 Then objDebug.WriteLine "GetFileLineCountSelect: vFileLines(" & nIndex & ")="  & vFileLines(nIndex) End If  
			nIndex = nIndex + 1
		End if
	Loop
	objDataFileName.Close
    GetFileLineCountInclude = nIndex
End Function
'-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function
'-----------------------------------------------------------------
'     Function GetDateFormat(nFormat)
'     nFormat: 1 = m/d/yyyy
'              2 = yyyy-mm-dd
'-----------------------------------------------------------------
Function GetDateFormat(MyDate, nFormat)
Dim strDate, strDay, strMonth, strYear
    If IsDate(MyDate) Then 
	    strDate = MyDate
	Else
	    strDate = Date()
		Call TrDebug("ERROR: Can't translate Date: " & MyDate, "", objDebug, 80 , 1, 1)
		Call TrDebug("CURRENT DATE WILL BE USED: " & Date(), "", objDebug, 80 , 1, 1)
	End If
    Select Case nFormat
	    Case 1
	        GetDateFormat = Month(strDate) & "/" & Day(strDate) & "/" & Year(strDate) 
		Case 2
		    If Year(strDate) <= 9 Then strYear = "0" & Year(strDate) Else strYear = Year(strDate)
			If Month(strDate) <= 9 Then strMonth = "0" & MOnth(strDate) Else strMonth = MOnth(strDate)
			If Day(strDate) <= 9 Then strDay = "0" & Day(strDate) Else strDay = Day(strDate)
			GetDateFormat = strYear  & "-" & strMonth & "-" & strDay
		Case Else 
		    GetDateFormat = Month(strDate) & "/" & Day(strDate) & "/" & Year(strDate) 
	End Select
End Function
'--------------------------------------------------------------
' Function returns a random intiger between min and max
'--------------------------------------------------------------
Function My_Random(min, max)
	Randomize
	My_Random = (Int((max-min+1)*Rnd+min))
End Function
' ----------------------------------------------------------------------------------------------
'   Function  TrDebug (strTitle, strString, objDebug)
'   nFormat: 
'	0 - As is
'	1 - Strach
'	2 - Center
' ----------------------------------------------------------------------------------------------
Function  TrDebug (strTitle, strString, objDebug, nChar, vFormat, nDebug)
Dim strLine
Dim nFormat
	If IsArray(vFormat) Then 
	    nFormat = vFormat(0)
		Set_A_Date = vFormat(1)
	Else 
	    nFormat = vFormat
		Set_A_Date = True
    End If	
	strLine = ""
	If nDebug <> 1 Then Exit Function End If
	If IsObject(objDebug) Then 
		Select Case nFormat
			Case 0
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) Else strLine = ""
				strLine = strLine & ":  " & strTitle
				strLine = strLIne & strString
				objDebug.WriteLine strLine
				
			Case 1
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) Else strLine = ""
				strLine = strLine & ":  " & strTitle
				If nChar - Len(strLine) - Len(strString) > 0 Then 
					strLine = strLine & Space(nChar - Len(strLine) - Len(strString)) & strString
				Else 
					strLine = strLine & " " & strString
				End If
				objDebug.WriteLine strLine
			Case 2
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				
				If nChar - Len(strLine & strTitle & strString) > 0 Then 
						strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
				Else 
						strLine = strLine & strTitle & " " & strString	
				End If
				objDebug.WriteLine strLine
			Case 3
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				For i = 0 to nChar - Len(strLine)
					strLIne = strLIne & "-"
				Next
				objDebug.WriteLine strLine
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				If nChar - 1 - Len(strLine & strTitle & strString) > 0 Then 
						strLine = strLine & Space(Int((nChar - 1 - Len(strLine & strTitle & strString))/2)) & strTitle & " " & strString			
				Else 
						strLine = strLine & strTitle & " " & strString	
				End If
				objDebug.WriteLine strLine
				If Set_A_Date Then strLine = GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  " Else strLine = ""
				For i = 0 to nChar - Len(strLine)
					strLine = strLine & "-"
				Next
				objDebug.WriteLine strLine
		End Select
	End If
End Function
