'----------------------------------------------------------------------------------
'            KIDS SCHILD MAIN LAUNCHER FORM SCRIPT VERSION 11.0
'----------------------------------------------------------------------------------
	Dim strDirectoryLCL
strDirectoryLCL = "C:\Users\All Users\Vandyke"
' KEEP STRDIRECTORY VARIABLE AT LINE 5 TO LET AUTOINSTALATION
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
'Const HttpTextColor1 = "#2C2A23"
'Const HttpTextColor2 = "#F0CD04"
Const HttpTextColor1 = "#292626"
Const HttpTextColor2 = "#F0BC1F"
Const HttpTextColor3 = "#EBEAF7"
Const HttpTextColor4 = "#A4A4A4"
Const HttpTextColor5 = "#Grey"
Const HttpBgColor1 = "Grey"
Const HttpBgColor2 = "#292626"
Const HttpBgColor3 = "#0404B4"
Const HttpBgColor4 = "#504E4E"
Const HttpBgColor5 = "#0D057F"
Const HttpBgColor6 = "#8B9091"
Const HttpBdColor1 = "Grey"
'	BackGroundColor = "blue"
'	ButtonColor = HttpBgColor5
'	MainTextColor = HttpTextColor3
'   InputBGColor = HttpBgColor5
Const fSZ_1 = 20
Const brdST_1 = "solid ; border-color: #5858FA; border-width: 1px"
Const OVER_TIME = 		"OVER_TIME"
Const GAME_OVER = 		"GAME_OVER"
Const NO_ACCOUNT = 		"NO_ACCOUNT"
Const ACTIVE = 			"ACTIVE"
Const SHUTTING_DOWN = 	"SHUTTING_DOWN"
Const DOWN = 			"DOWN"
Const INACTIVE = 		"INACTIVE"
Const OFF_LINE =		"OFF_LINE"
Const ON_LINE = 		"ON_LINE"
Const UNKNOWN = 		"UNKNOWN"
Const NO_INTERNET = 	"NO_INTERNET"
Const SERVER_UP =		"UP"
Const AUTO_TIMER = 		"AUTO"
Const NO_LOGON = 		"NO_LOGON"
Const MAX_TIMER = 		20
Const PARENTS = 		"1"
Const IE_PAUSE = 		70
Const MAIN_TITLE =      "KidsShield Launcher"
Const DEVELOPER_HOST = "K-MESON-W7-32"
' - Global Variables to Access Server
Dim strNetworkStatus, strServerStatus
Dim strGatewayIP, strServerIP, strFileConnect
' - Global Varibles for Local Desktop
Dim strWinUser, strHost
' Other Global Variables
Dim nResult
Dim strMonthMaxFileName, strFileAccount, strFileDevice, strFileString, strFileUnattend, strRestartFile
Dim strFileSession, strFileWinUserTmp, strFileWinUser, strFileScreenUser
Dim strSrvDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke, strCRTexe
Dim strDeviceID, strLclDeviceID
Dim strAccountID, strOwnerOld, strOwnerID
Dim nDebug, ShowDebug, nInfo
Dim vConnect
Dim nDevice, nAccount, nActivity, nSession, nSessionTmp, nInventory
Dim vDevice, vAccount, vSession, vMsg(20), vvMsg(20,3), vInventory
Dim vLine, vIE_Scale
Dim vAllUsersSchedulers(20,10)
Dim objFSO, objEnvar, objDebug, objSession, objSessionTmp, objShell
Dim LauncherPID, IE_Window_Title, Restart
strFileSession = "sessions.txt"
strFileSessionTmp = "sessions_tmp.dat"
strSrvDirectory = "\\MEDIA\_PublicFolder\KidsSchild"
strDirectoryWork = "C:\KidsSchild\DVLP"
strDirectoryUpdate = "\\HIGS\Install_My\Tools_Networking\KidsSchild"
strDirectoryVandyke = "C:\Program Files"
strCRTexe = "\SecureCRT.exe"""
strFileInventory = "inventory.ini"
strFileWinUserTmp = "winusers_tmp.dat"
strFileConnect = "connectivity_tmp.dat"
strFileScreenUser = "screenuser_tmp.dat"
strVersion = "None"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objEnvar = WScript.CreateObject("WScript.Shell")
Set objSession = objFSO.OpenTextFile(strDirectoryLCL & "\" & "sessions.txt")
Set objShell = WScript.CreateObject("WScript.Shell")
strDirectoryTmp = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%") & "\AppData\Local\Temp"
strHost = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
nDebug = 0
nInfo = 1
ShowDebug = False
Restart = False
Main()
If IsObject(objDebug) Then objDebug.Close 
set objDebug = Nothing
set objFSO = Nothing
set objEnvar = Nothing
Set objShell = Nothing
If Restart Then Call LaunchScript("LAUNCHER",vInventory, nInventory, "" , False) : End If
Sub Main()
'-----------------------------------------------------------------
'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
'-------------------------------------------------------------------------------------------
'  OPEN LOG FILE AND CHECK IF ANOTHER INSTANCE OF LAUNCHER IS ALREADY RUNNING
'-------------------------------------------------------------------------------------------
	If Not objFSO.FolderExists(strDirectoryLCL & "\Log") Then 
			objFSO.CreateFolder(strDirectoryLCL & "\Log") 
	End If
	
	On Error Resume Next
	strLogFile = strDirectoryLCL & "\Log\" & "debug-launcher-" & GetLogonUser(0) & ".log"
	Set objDebug = objFSO.OpenTextFile(strLogFile,ForWriting,True)
	Select Case Err.Number
		Case 0
		Case 70
		'	vvMsg(0,0) = "LAUNCHER IS ALREADY RUNNING" 					: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
		'	vvMsg(1,0) = "Exit . . ."									: vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2 
 		'	Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,2)
			Call IE_GetPID(LauncherPID, MAIN_TITLE & " - " & IE_Window_Title, 0)
			Call FocusToParentWindow(LauncherPID)			
		Exit Sub
		Case Else 
			vvMsg(0,0) = "CAN'T START LAUNCHER" 						: vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Exit . . ."									: vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor1
 			Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,2)
			Exit Sub
	End Select
	On Error goto 0
'-----------------------------------------------------------------
'  Read default parameters
'-----------------------------------------------------------------
	If objFSO.FileExists(strDirectoryLCL & "\" & strFileSession) Then 
		nSession = GetFileLineCountSelect(strDirectoryLCL & "\" & strFileSession, vSession,"#","","NULL",nDebug)
        strSrvDirectory = vSession(1)
		If Right(strSrvDirectory,1) = "\" Then 
			strSrvDirectory = Left(strSrvDirectory,Len(strSrvDirectory) - 1)
			If nDebug = 1 Then objDebug.WriteLine FormatDateTime(Date(), 0) & " " & FormatDateTime(Time(), 3) & ": Remote Server Folder: " & strSrvDirectory End If
		End If
		If nSession > 0 Then	strDirectoryVandyke = vSession(0)	End If	' - Vandyke folder to run SecureCRT from 
		If nSession > 3 Then	strDirectoryWork = vSession(3)		End If	' - Work directory scripts are installed to 
		If nSession > 4 Then    strDirectoryUpdate = vSession(4)	End If	' - Source directory to take updates from 
		If nSession > 5 Then    strVersion = vSession(5)			End If	' - Current version of the Package/Launcer
		If nSession > 6 Then	strLclDeviceID = vSession(6)    	End If	' - Name of the Local PC is stored in session.txt line 7
		If nSession > 7 Then	strOwnerID = vSession(7)    		End If	' - Name of the Owner of the Local PC is stored in session.txt line 7
		If nSession > 8 Then	strServerIP = vSession(8)  			End If	' - Server IP address is stored in session.txt line 9
		If nSession >=10 Then	strGatewayIP = vSession(9) 			End If	' - Gateway IP address is stored in session.txt line 10
	End If
	GwIP = GetDefaultGwWMI(".", nDebug)
	IF GwIP <> "0" Then strGatewayIP = GwIP
	strCRTexe = """" & strDirectoryVandyke & strCRTexe
	strFileAccount = strSrvDirectory & "\" & "accounts.dat"
	strFileDevice = strSrvDirectory & "\" & "devices.dat"
	strFileString = strSrvDirectory & "\" & "string.dat"
	strMonthMaxFileName = strSrvDirectory & "\" & "month-max.dat"
    strFileUnattend = strDirectoryLCL & "\Temp\Unattend.dat"	
	strRestartFile = strDirectoryLCL & "\Temp\restart.dat"	
	strWinUtilsFolder = strDirectoryWork & "\Bin"
	strFileDeviceCatalog = strSrvDirectory & "\data\devices_catalog.dat"
	If ShowDebug Then 
	    Set objBat = objFSO.OpenTextFile(Replace(strLogFile,"log","bat"),ForWriting,True)
		objBat.WriteLine strWinUtilsFolder & "\tail.exe -f " &  """" & strLogFile & """ -n 40"
		objBat.close
		Set objBat = Nothing  
		objShell.run """" & Replace(strLogFile,"log","bat") & """",1
	End If
	'-----------------------------------------------------------------
    '  GET SCREEN RESOLUTION
    '-----------------------------------------------------------------
	Call WriteScreenResolution(vIE_Scale, 0)
	'--------------------------------------------
	'   SETUP DIRECTORY ARRAY
	'--------------------------------------------
	Dim vDir(10)
	vDir(0) = strDirectoryLCL
	vDir(1) = strSrvDirectory
	vDir(2) = strDirectoryWork
	vDir(3) = strDirectoryUpdate
    vDir(4) = strDirectoryVandyke
	'-------------------------------------------------------------------------------------------
	'			GET NAME OF THE LOCALY LOGED-ON WINDOWS USER. GET DESKTOP CONTEXT
	'-------------------------------------------------------------------------------------------
	strWinUser = GetScreenUserSYS()
	If strWinUser <> GetScreenUser(strDirectoryLCL & "\Temp\" & strFileScreenUser, nDubug) Then 
		Call SetScreenUser(strDirectoryLCL & "\Temp\" & strFileScreenUser, strWinUser,nDebug) 
	End If
	'-------------------------------------------------------------------------------------------
	'           CHECK LOCAL PC CONNECTIVITY TO THE NETWORK AND SERVER AVAILABILITY
	'-------------------------------------------------------------------------------------------
	strNetworkStatus = OFF_LINE
	strServerStatus = OFF_LINE	
    Call GetConnectStatus (strNetworkStatus, strServerStatus, strServerIP, strGatewayIP, strDirectoryLCL & "\Temp\" & strFileConnect, 1000, 1 ,False, nDebug)
	If strServerStatus = ON_LINE Then  Call GetServerStatus (strFileAccount, strNetworkStatus, strServerStatus, strDirectoryLCL & "\Temp\" & strFileConnect, 1000, 1 ,False, nDebug) End if
	Call TrDebug ("CHECK CONNECTIVITY TO SERVER:", "OK", objDebug, MAX_LEN, 1, nDebug)
	If strNetworkStatus <> ON_LINE Then Call TrDebug ("CONNECTIVITY TO NETWORK: ", strNetworkStatus, objDebug, MAX_LEN, 1, nInfo)
	If strServerStatus <> SERVER_UP Then Call TrDebug ("CONNECTIVITY TO SERVER:",strServerStatus , objDebug, MAX_LEN, 1, nInfo)	
	'-------------------------------------------------------------------------------------------
	'           CHECK LOCAL CLNT AGENT IS UP AND RUNNING
	'-------------------------------------------------------------------------------------------
	Dim objFileConnect
	If objFSO.FileExists(strDirectoryLCL & "\Temp\" & strFileConnect) Then
		Set objFileConnect = objFSO.GetFile(strDirectoryLCL & "\Temp\" & strFileConnect)
		If DateDiff("n", objFileConnect.DateLastModified, Date() & " " & Time()) => 2 and GetHostNameSYS() <> DEVELOPER_HOST Then strServerStatus = ON_LINE
	Else 
		strServerStatus = OFF_LINE	
	End If
'	If objFSO.FileExists(strDirectoryLCL & "\Temp\" & strFileConnect) Then 
'		nConnect = GetFileLineCountSelect(strDirectoryLCL & "\Temp\" & strFileConnect, vConnect,"#","","NULL",nDebug)
'		strNetworkStatus = Split(vConnect(0),",")(1)
'		strServerStatus = Split(vConnect(1),",")(1)
'	End If
'--------------------------------------------------------------------------------
'               READ LIST OF CLIENT SCRIPTS FROM INVENTORY FILE
'--------------------------------------------------------------------------------
	strFile = strDirectoryWork & "\" & strFileInventory
	nInventory = GetFileLineCountByGroup(strFile, vInventory,"Server","Client","",0)
'--------------------------------------------------------------------------------
'               CHECK IF THE AutoLogin ACCOUNT BELONG TO ADMINISTRATOR
'--------------------------------------------------------------------------------
	If strServerStatus = SERVER_UP Then  
		Call GetFileLineCountSelect(strFileAccount, vAccount,"#","$","", 0)
		nAccount = GetObjectLineNumber(vAccount,UBound(vAccount),strOwnerID, True)
		If nAccount > 0 Then 
			If Split(vAccount(nAccount - 1),",")(3) = PARENTS Then strOwnerID = NO_LOGON End If
		End If
		If nAccount = 0 Then strOwnerID = NO_LOGON End If
	End If
	nAdmin = "0"
'-------------------------------------------------------------------------------
'   WRITE NAME OF THE SCREEN USER TO SERVER
'--------------------------------------------------------------------------------
    If strServerStatus = SERVER_UP Then  
	    strWinUser = GetScreenUserSYS()
		If Not WriteScreenUserToServer(strSrvDirectory & "\logon-" & strLclDeviceID & ".dat", strWinUser, nDebug) Then 
			Call TrDebug ("SAVING NAME OF THE DESKTOP LOGON TO SERVER: " & strWinUser, "ERROR", objDebug, MAX_LEN, 1, 1)
		Else
			Call TrDebug ("SAVING NAME OF THE DESKTOP LOGON TO SERVER: " & strWinUser, "OK", objDebug, MAX_LEN, 1, 1)
		End If
	End If
'--------------------------------------------------------------------------------
'   START MAIN PROGRAM
'--------------------------------------------------------------------------------	
	Do
		'------------------------------------------------------------------
		'	IE_LAUNCHER RETURNS:
		'	False 	Press OK
		'   1		UPDATE OR RECONNECT
		'------------------------------------------------------------------
 		nResult = IE_LAUNCHER (vIE_Scale, nAdmin, strOwnerID, strVersion, vInventory, vAccount, vDir, strFileAccount, 0)
		Select Case nResult 
			Case False
				'------ Enable Auto Logout for Parent's account ----- ' 
				If nAdmin = PARENTS Then Call WriteStrToFile(strDirectoryLCL & "\" & strFileSession, NO_LOGON, 8, 1, nDebug) End If
				'-----------------------------------------------------------
				'   CLEAN LAUNCHER TEMP FILES
				'-----------------------------------------------------------
				strFolder = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
				Call CleanFolder(strFolder, "run-", "bat", 60, 0)
				Call CleanFolder(strFolder, "starting_my_vm", "dat",60,0)
				Call CleanFolder(strFolder, "svc-", "dat",60,0)	
				strFolder = strDirectoryLCL & "\Temp"
				Call CleanFolder(strFolder, "unattend", "dat",2,0)
				Call CleanFolder(strFolder, "telnet", "dat",30,0)
				Call CleanFolder(strFolder, "week", "dat",8 * 24 * 60,0)
				Exit Do
			Case 1
				Restart = True                        						
				Exit Do
		End Select
	Loop
End Sub
'##############################################################################
'      Function Create Launcher Form
'##############################################################################
 Function IE_LAUNCHER (vIE_Scale, ByRef nAdmin, ByRef strOwnerID, strVersion, ByRef vInventory, Byref vAccount, vDir, strFileAccount, nDebug)
    Dim strDirectoryLCL, strSrvDirectory, strDirectoryWork, strDirectoryUpdate, strDirectoryVandyke, SettingsFigure, nMenuButtonX, nMenuButtonY, nPressSettings, nTab
    Dim strAccountID, nInventory, MenuH
	Dim NewPasswd, OldPasswd
	Dim g_objIE
    Dim g_objShell
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim strHandler, strLogin
	Dim vUpdate, strUpdate, nUpdate, nVersion, strSERVER_BG, strSERVER_C, strUPDATE_BG, strUPDATE_C
	Dim IE_Menu_Bar
	Dim  IE_Border
	Dim strPID
    Dim IE_Window_Title
	Dim strBlock
	Dim strConfigFile, strBkpConfigFile, strNewConfigFile
	nInventory = UBound(vInventory)
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	Set g_objShell = WScript.CreateObject("WScript.Shell")	
	' - FOLDER LIST
	strDirectoryLCL = vDir(0)
	strSrvDirectory = vDir(1)
    strDirectoryWork = vDir(2)
	strDirectoryUpdate = vDir(3)
    strDirectoryVandyke = vDir(4)
	' - Get Current Status
	'----------------------------------------------------------------
	'   Title Figure
	'----------------------------------------------------------------
	TitleFigure = strDirectoryWork & "\bin\launcher_title.jpg"	
    BottomFigure = strDirectoryWork & "\bin\launcher_bottom.jpg"
    strConfigFile = strSrvDirectory & "\config.ini"
	strNewConfigFile = strSrvDirectory & "\config.new"
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  g_objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080	
	Dim nTimeOut, Tsys0
	Const IE_PAUSE = 70
	Const TIME_OUT = 10
	Tsys0 = Date() & " " & Time()
	IE_LAUNCHER = False
	strSERVER_BG = HttpBgColor3
	strSERVER_C = HttpTextColor2
	strUPDATE_BG = HttpBgColor3
	strUPDATE_C = HttpTextColor2
    Call Set_IE_obj (g_objIE)
	nButtonX = Round(150 * nRatioX,0)
	nButtonY = Round(150 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 50 then nButtonY = 50 End If
	LoginTitleW = nButtonX * 3
	LoginTitleH = Round(40 * nRatioY,0) * 2
	WindowW = IE_Border + LoginTitleW
	WindowH = IE_Menu_Bar + nButtonY * 3 + 2 * LoginTitleH
	MenuH = WindowH - IE_Menu_Bar
	nTab = 20
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)	
	nFontSize_24 = Round(24 * nRatioY,0)
	g_objIE.Offline = True		
	g_objIE.navigate "about:blank"
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "None"
	g_objIE.Document.body.Style.background = HttpBgColor3
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	g_objIE.document.Title = MAIN_TITLE
	SettingsFigure = strDirectoryWork & "\bin\settings-icon.png"
	strHTMLBody = ""
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" background=""" & TitleFigure & """ style="" position: absolute; left: 0px; top: 0px;" &_
		"background-repeat: no-repeat; background-position: 50% 50%; background-size: " & LoginTitleW & "px " & LoginTitleH & "px;" &_
		" border-collapse: collapse; border-style: none; border-width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: transparent; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td background=""" & SettingsFigure & """ style="" background-repeat: no-repeat; background-position: 10% 50%; background-size: 40px 40px;" &_
			    " border-style: none; background-color: transparent;""" &_
				"align=""left"" valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
				"<button style='position: relative; left: 10px; background-color: transparent; border-style: None; width:" &_
				"40;height:40;'" &_
				" id='SETTINGS' name='SETTINGS' onclick=document.all('ButtonHandler').value='SETTINGS';></button>" & _	
			"</td>" &_
		    "<td style="" border-style: none; background-color: transparent;""valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
		        "<input name='GreetingTitle' id='Greeting' size='10' maxlength='10' style=""text-align: center; border-style: None; font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
				";background-color: transparent; font-weight: normal;""></input>" & _
			"</td>" &_
			"<td style="" border-style: none; background-color: transparent;""valign=""middle"" align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/3) & """>" & _
				"<button style='border-style: none; background-color: transparent" &_
				"; font-size: " & nFontSize_14 &_
				"; font-weight: Normal" &_
				"; color: " & HttpTextColor3 & "; width: " & Int(LoginTitleW/3) & "px ;height: " & LoginTitleH & "px"  &_
    			"'id='LOGIN' name='LOGIN' AccessKey='L' onclick=document.all('ButtonHandler').value='LOGIN'></button>" &_		
     		"</td>" & _
		"</tr></tbody></table>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: 0" &_ 
				"px; bottom: " & 2 * nButtonY + LoginTitleH & "px' id='START' name='START' AccessKey='S' " &_
				"onclick=document.all('ButtonHandler').value='START'" &_
				"><u>S</u>TART</button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nButtonX &_
				"px; bottom: " & 2 * nButtonY + LoginTitleH & "px' name='STOP' AccessKey='T' onclick=document.all('ButtonHandler').value='STOP';>STO<u>P</u></button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & 2 * nButtonX &_
				"px; bottom: " & 2 * nButtonY + LoginTitleH & "px' name='TODO_LIST' AccessKey='T' onclick=document.all('ButtonHandler').value='TODO_LIST';><u>T</u>O-DO LIST</button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: 0" &_ 
				"px; bottom: " & LoginTitleH & "px' name='ACTIVE' AccessKey='V' onclick=document.all('ButtonHandler').value='ACTIVE';>ACTI<u>V</u>E</button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nButtonX &_
				"px; bottom: " & LoginTitleH & "px' name='SHOW_ALL' AccessKey='A' onclick=document.all('ButtonHandler').value='ALL';><u>A</u>LL</button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & strUPDATE_BG &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & strUPDATE_C & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & 2 * nButtonX &_
				"px; bottom: " & LoginTitleH & "px' id='UPDATE' name='UPDATE' AccessKey='U' onclick=document.all('ButtonHandler').value='UPDATE';></button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & strSERVER_BG &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & strSERVER_C & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: 0" &_ 
				"px; bottom: " & nButtonY + LoginTitleH  & "px' id='SERVER' name='SERVER' AccessKey=E' onclick=document.all('ButtonHandler').value='SERVER';></button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & nButtonX &_
				"px; bottom: " & nButtonY + LoginTitleH   & "px' name='PARENT' AccessKey='P' onclick=document.all('ButtonHandler').value='PARENT';><u>P</u>ARENTS</button>"
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: " & brdST_1 & "; background-color: " & HttpBgColor3 &_
				"; font-size: " & fSZ_1 &_
				"; font-weight: bold" &_
				"; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & 2 * nButtonX &_
				"px; bottom: " & nButtonY + LoginTitleH & "px' name='TOKENS' AccessKey='T' onclick=document.all('ButtonHandler').value='BOOKS';>T<u>O</u>KENS</button>"
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" background=""" & BottomFigure & """ style="" position: absolute; left: 0px; bottom: 0px;" &_
		" border-collapse: collapse; border-style: none; border-width: 1px; border-color: " & HttpBgColor5 &_
		"background-repeat: no-repeat; background-position: 50% 50%; background-size: " & LoginTitleW & "px " & LoginTitleH & "px;" &_		
		"; background-color: " & HttpBgColor5 & "; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td style="" border-style: none; background-color: transparent;""valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & 2 * LoginTitleW/3 & """>" & _
				"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
				"; font-weight: normal;font-style: italic;"">&nbsp;&nbsp;Version " & strVersion & "</span> </p>" & _
			"</td>" &_
			"<td style="" border-style: none; background-color: transparent;""valign=""middle"" align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW/3 & """>" & _
     		"</td>" & _
		"</tr></tbody></table>"				
    strHTMLBody = strHTMLBody &_				
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
	'-----------------------------------------------------------------
	' SETTINGS MENU
	'-----------------------------------------------------------------
	nMenuButtonX = 200 - nTab
	nMenuButtonY = Int(LoginTitleH/2)
	strHTMLBody = strHTMLBody &_	
	 "<div id='divSettings' name='divSettings' style='color: " & HttpTextColor3 & " ;background-color:" & HttpBgColor5 & "; width: 200px; height: " & MenuH - 2 * LoginTitleH &_
	 "px; position: absolute; left: -200px; top: " & LoginTitleH & "px;'>" &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: relative; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: Transparent" &_
		"; width: " & Int(LoginTitleW/4) & "px;"">" & _
			"<tbody>" & _
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/2) & """ width=""" & Int(LoginTitleW/4) & """>" & _
					"</td>"&_
				"</tr>" &_
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/2) & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' name='CHANGE_PASS' onclick=document.all('ButtonHandler').value='CHANGE_PASS';>Change Password</button>" & _	
					"</td>"&_
				"</tr>" &_
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/2) & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' id='SET_KID_PASS' name='SET_KID_PASS' onclick=document.all('ButtonHandler').value='SET_KID_PASS'>Set Kids Password</button>" & _	
					"</td>"&_
				"</tr>" &_
    			"<tr>" &_
    				"<td style=""border-style: None; background-color: Transparent;""align=""right"" class=""oa1"" height=""" & Int(LoginTitleH/2) & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor5 & "; color: " & HttpTextColor3 & "; width:" &_
						nMenuButtonX & ";height:" & nMenuButtonY & "; font-size: " & nFontSize_12 & ".0pt;" &_
						"px; ' id='SIGN_CONFIG' name='SIGN_CONFIG' onclick=document.all('ButtonHandler').value='SIGN_CONFIG'>Sign Config</button>" & _	
					"</td>"&_
				"</tr>" &_								
		    "</tbody></table>" &_
	 "</div>"
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.Visible = False
	'----------------------------------------------------
	'    SET INITIAL VALUES
	'----------------------------------------------------
	If objFSO.FileExists(strRestartFile) Then 
	    Call GetFileLineCountSelect(strRestartFile, vRestart,"#","","NULL",0)
		Select Case Split(vRestart(0),":")(0)
		    Case "SUCCESS"
		        objFSO.DeleteFile strRestartFile, True
		    Case Else
               If strServerStatus = SERVER_UP Then strServerStatus = ON_LINE
        End Select
    End If	
	If strOwnerID = "NO_LOGON" Then 
		strLine = "" 
		strLogOnButton = "Sign In"
	else 
		strLine = "Hi, " & strOwnerID & "!"
		strLogOnButton = "Sign Out"
	End if
    If nAdmin <> PARENTS Then    
		   g_objIE.Document.getElementById("SET_KID_PASS").disabled = True 
		   g_objIE.Document.getElementById("SET_KID_PASS").style.color = "Grey"
		   g_objIE.Document.getElementById("SIGN_CONFIG").disabled = True 
		   g_objIE.Document.getElementById("SIGN_CONFIG").style.color = "Grey"		   
	Else 
		   g_objIE.Document.getElementById("SET_KID_PASS").disabled = False 
		   g_objIE.Document.getElementById("SET_KID_PASS").style.color = HttpTextColor3
		   g_objIE.Document.getElementById("SIGN_CONFIG").disabled = False
		   g_objIE.Document.getElementById("SIGN_CONFIG").style.color = HttpTextColor3
	End If
    g_objIE.Document.getElementById("UPDATE").style.backgroundcolor = HttpBgColor3
	g_objIE.Document.getElementById("UPDATE").style.color = HttpTextColor2
	g_objIE.Document.getElementById("UPDATE").innerHTML = "UPDATE"
	g_objIE.Document.getElementById("UPDATE").disabled = True
	g_objIE.Document.getElementById("LOGIN").innerHTML = strLogOnButton
	g_objIE.Document.getElementById("LOGIN").style.fontfamily = "arial narrow"	
    g_objIE.Document.getElementById("LOGIN").style.fontsize = "18"
    g_objIE.Document.getElementById("LOGIN").style.fontweight = "normal"	
    g_objIE.Document.getElementById("LOGIN").style.color = HttpTextColor3	
    g_objIE.Document.getElementById("Greeting").style.fontsize = "18"
    g_objIE.Document.getElementById("Greeting").style.fontweight = "normal"	
    g_objIE.Document.getElementById("Greeting").style.color = HttpTextColor3		
	g_objIE.Document.All("GreetingTitle").value = strLine	
	If strNetworkStatus = OFF_LINE or strServerStatus = OFF_LINE Then 
		g_objIE.Document.getElementById("LOGIN").disabled = True		
		g_objIE.Document.getElementById("SERVER").style.backgroundcolor = "blue" 
		g_objIE.Document.getElementById("SERVER").style.color = "red"
	    g_objIE.Document.getElementById("SERVER").innerHTML = "NOT<br>CONNECTED"		
	End If
	If strServerStatus = ON_LINE Then 
		g_objIE.Document.getElementById("LOGIN").disabled = True
		g_objIE.Document.getElementById("SERVER").style.backgroundcolor = "Blue" 
		g_objIE.Document.getElementById("SERVER").style.color = "red"
	    g_objIE.Document.getElementById("SERVER").innerHTML = "NOT<br>CONNECTED"		
	End If
	If strServerStatus = SERVER_UP Then 
	    strFileUpdate = strDirectoryUpdate & "\" & "updatelist.ini"
		nUpdate = GetFileLineCountSelect(strFileUpdate, vUpdate, "", "NULL","NULL", 0)
		If nUpdate > 0 Then 
			If vUpdate( nUpdate - 1 ) <> strVersion Then 
				g_objIE.Document.getElementById("UPDATE").style.backgroundcolor = "blue"
				g_objIE.Document.getElementById("UPDATE").style.color = HttpTextColor3
				g_objIE.Document.getElementById("UPDATE").innerHTML = "NEW <br>v." & vUpdate( nUpdate - 1 )
			End If
		End If
		g_objIE.Document.getElementById("SERVER").style.backgroundcolor = HttpBgColor3 
		g_objIE.Document.getElementById("SERVER").style.color =  HttpTextColor2		
		g_objIE.Document.getElementById("SERVER").innerHTML = "CONNECTED"
		g_objIE.Document.getElementById("LOGIN").disabled = False
		g_objIE.Document.getElementById("UPDATE").disabled = False
	End If
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	Call IE_Unhide(g_objIE)
	WScript.sleep 100
	' Call IE_GetPID(strPID, MAIN_TITLE & " - " & IE_Window_Title, nDebug)
	' g_objShell.AppActivate strPID
	strPID = MD5Hash("password")
	Do
		'---------------------------------------------------------------
		'   Check TimeOut Value
		'---------------------------------------------------------------
		If DateDiff("s", Tsys0, Date() & " " & Time()) > TIME_OUT * 60 Then
			Select Case nAdmin
				Case "1" ' - Then Log Out from Amin Account
					Tsys0 = Date() & " " & Time()
					g_objIE.Document.All("ButtonHandler").Value = "LOGIN"
					Call TrDebug ("AUTO LOGOUT: " & strOwnerID & " -> NO_LOGIN", "", objDebug, MAX_LEN, 1, 1)
				Case "0"
					Tsys0 = Date() & " " & Time()
			End Select
			' - Update ScreenUser 
		    Call GetConnectStatus (strNetworkStatus, strServerStatus, strServerIP, strGatewayIP, strDirectoryLCL & "\Temp\" & strFileConnect, 1000, 1 ,False, nDebug)
	        If strServerStatus = ON_LINE Then  Call GetServerStatus (strSrvDirectory & "\" & "accounts.dat", strNetworkStatus, strServerStatus, strDirectoryLCL & "\Temp\" & strFileConnect, 1000, 1 ,False, nDebug) End if			
			If strNetworkStatus = OFF_LINE or strServerStatus = OFF_LINE Then 
				g_objIE.Document.getElementById("LOGIN").disabled = True		
				g_objIE.Document.getElementById("SERVER").style.backgroundcolor = "blue" 
				g_objIE.Document.getElementById("SERVER").style.color = "red"
				g_objIE.Document.getElementById("SERVER").innerHTML = "NOT<br>CONNECTED"		
			End If
			If strServerStatus = ON_LINE Then 
				g_objIE.Document.getElementById("LOGIN").disabled = True
				g_objIE.Document.getElementById("SERVER").style.backgroundcolor = "Blue" 
				g_objIE.Document.getElementById("SERVER").style.color = "yellow"
				g_objIE.Document.getElementById("SERVER").innerHTML = "NOT<br>CONNECTED"		
			End If
			If strServerStatus = SERVER_UP Then 
				' - Check For Updates on Server
				strFileUpdate = strDirectoryUpdate & "\" & "updatelist.ini"
				nUpdate = GetFileLineCountSelect(strFileUpdate, vUpdate, "", "NULL","NULL", 0)
				If nUpdate > 0 Then 
					If vUpdate( nUpdate - 1 ) <> strVersion Then 
						g_objIE.Document.getElementById("UPDATE").style.backgroundcolor = "blue"
						g_objIE.Document.getElementById("UPDATE").style.color = HttpTextColor3
						g_objIE.Document.getElementById("UPDATE").innerHTML = "NEW <br>v." & vUpdate( nUpdate - 1 )
					End If
				End If
				g_objIE.Document.getElementById("SERVER").style.backgroundcolor = HttpBgColor3 
				g_objIE.Document.getElementById("SERVER").style.color =  HttpTextColor2		
				g_objIE.Document.getElementById("SERVER").innerHTML = "CONNECTED"
				g_objIE.Document.getElementById("LOGIN").disabled = False
				g_objIE.Document.getElementById("UPDATE").disabled = False
			End If			
		End If
		On Error Resume Next
		Err.Clear
		strHandler = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		If g_objIE.width <> WindowW Then g_objIE.width = WindowW End If
		If g_objIE.height <> WindowH Then g_objIE.height = WindowH End If
		On Error Goto 0
		Select Case strHandler
			Case "PARENT"    
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
				Do
					If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" Then 
						vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "TO START NEW SESSION"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
						vvMsg(2,0) = "Sign in . . . "			    		: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 3)
						Exit Do
					End If 
					If nAdmin <> PARENTS Then 
						vvMsg(0,0) = "RESTRICTED ACCESS! "		: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor2
						vvMsg(1,0) = "YOU NEED TO LOGIN AS  "	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
						vvMsg(2,0) = "PARENT FIRST."			: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor1		
						vvMsg(3,0) = "Try again..."				: vvMsg(3,1) = "bold" 	: vvMsg(3,2) = HttpTextColor2			
						Call IE_MSG(g_objIE,vIE_Scale, "Restricted Access",vvMsg,4)
						Exit Do
					End If 
					Call IE_Hide(g_objIE)					
					nResult = IE_PromptForAdminAction(vIE_Scale, vDir, strOwnerID, strAccountID, vAccount, nDebug)
					Select Case nResult
                        Case 0
					        Call IE_Unhide(g_objIE)											  
					    Case 1
						    g_objIE.Document.All("ButtonHandler").Value = "MNG_KIDS_TASK"
                        Case 2
						    g_objIE.Document.All("ButtonHandler").Value = "MNG_KIDS_TODO"
                        Case 3
						    g_objIE.Document.All("ButtonHandler").Value = "MNG_KIDS_STATUS"						
					End Select
					Exit Do
				Loop
			Case "SIGN_CONFIG"
			        g_objIE.Document.All("ButtonHandler").Value = "None"
					' - Close Settings Panel
					For MenuX = 200 to 0 step - 4
						g_objIE.document.getElementById("divSettings").style.left = (MenuX - 200) & "px"
						'WScript.Sleep 1
					 Next
					nPressSettings = nPressSettings + 1
                    Call IE_Hide(g_objIE)
					Do
						If objFSO.FileExists(strNewConfigFile) Then 
						    Call SetFileSignature(strNewConfigFile)
							Call FileHidden(strConfigFile, False)
							If Not objFSO.FolderExists(strSrvDirectory & "\Backup") Then objFSO.CreateFolder strSrvDirectory & "\Backup"
							strBkpConfigFile = strSrvDirectory & "\Backup\config-" & DateDiff("n",DateSerial(2015,1,1),Date() & " " & Time()) & ".ini"
							objFSO.CopyFile strConfigFile, strBkpConfigFile, True
							objFSO.DeleteFile strConfigFile, True
							objFSO.CopyFile strNewConfigFile, strConfigFile, True
							objFSO.DeleteFile strNewConfigFile, True
							Call FileHidden(strConfigFile, True)
							Call FileHidden(strNewConfigFile, True)
							Call FileHidden(strBkpConfigFile, True)
							vvMsg(0,0) = "NEW CONFIGURATION FILE SIGNED" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
							vvMsg(1,0) = "SUCCESSFULLY"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2
							Call IE_MSG(g_objIE,vIE_Scale, "Task Completed", vvMsg, 2)							
							Exit Do
						Else
							vvMsg(0,0) = "CAN'T FIND A CANDIDATE" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
							vvMsg(1,0) = "CONFIGURATION FILE"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
							vvMsg(2,0) = ".\config.new "			    		: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
							Call IE_MSG(g_objIE,vIE_Scale, "Can't find file", vvMsg, 3)
							Exit Do
						End if 
						Exit Do
					Loop
			Case "LOGIN"
			        g_objIE.Document.All("ButtonHandler").Value = "None"
					Do
					    Select Case g_objIE.Document.getElementById("LOGIN").innerHTML
						    Case "Sign In"
	    					    Call IE_Hide(g_objIE)
								If strOwnerID = "NO_LOGON" Then strLogin = "" Else strLogin = strOwnerID
								vvMsg(0,0) = "INTRODUCE YOURSELF" : vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor5
     							If  Not IE_PromptLoginPassword (g_objIE,vIE_Scale, vvMsg, 1, strLogin, Passwd,False, nDebug) Then Call IE_Unhide(g_objIE) :  Exit Do : End If
								PasswdMD5 = MD5Hash(Lcase(strLogin) & Passwd)
								strLine = Date() & " " & Time() & ": Login: " & strLogin & " Password: " & PasswdMD5
							    ' MsgBox Lcase(strLogin) & ", " & Passwd & " --> " & PasswdMD5
								If strLogin = "" Then Call IE_Unhide(g_objIE) :  Exit Do : End If
								nAcctIndex = GetObjectLineNumber(vAccount,UBound(vAccount),strLogin, False) - 1
                                If  nAcctIndex < 0 Then
									vvMsg(0,0) = "USER DOESN'T EXIST"   :  vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
									vvMsg(1,0) = "TRY AGAIN . . ."      :  vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2
									Call IE_MSG(g_objIE,vIE_Scale, "Wrong Password", vvMsg, 2)
									strOwnerID = strLogin
									Exit Do
								End If
								On Error Resume Next
								Err.Clear
								strOwnerID = Split(vAccount(nAcctIndex),",")(0)
								If PasswdMD5 <> Split(vAccount(nAcctIndex),",")(2) Then 
									Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", False)
									Call WriteStrToFile(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", strLine & "   FAILED", 100, 4, nDebug)
									Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", True)
									If Err.Number > 0 Then 
									    MsgBox "Error: Line 705: " & Err.Description & chr(13) &_
                        									    "strOwnerID=" & strOwnerID & chr(13) &_
																"AcctIndex=" & nAcctIndex & chr(13) &_
																"vAccount(nAcctIndex)=" & vAccount(nAcctIndex) & chr(13)
									    Exit Do
									End If 
									vvMsg(0,0) = "WRONG PASSWORD!"          :  vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
									vvMsg(1,0) = "TRY AGAIN . . ."      :  vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2
									Call IE_MSG(g_objIE,vIE_Scale, "Wrong Password", vvMsg, 2)
									On Error Goto 0
									Exit Do
								End If
								Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", False)
								Call WriteStrToFile(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", strLine & "   OK", 100, 4, nDebug)
								Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", True)
								On Error Goto 0
                                g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign Out"
								g_objIE.Document.All("GreetingTitle").value = "Hi, " & strOwnerID & "!"
                                Call WriteStrToFile(strDirectoryLCL & "\" & strFileSession, strOwnerID, 8, 1, nDebug)
								' - GET AN ADMIN STATUS OF THE USER: 1 = PARENTS (admin) 
								Call GetAccountStatus(vAccount,nAcctIndex,nAdmin,strBlock,nDebug)
								Call TrDebug ("Account: " & strAccountID & "Parent: " & nAdmin & " Blocked: " & strBlock,"", objDebug, MAX_LEN, 1, nDebug)								
								' nAdmin = Split(vAccount(nAcctIndex),",")(3)
								Call IE_Unhide(g_objIE)
								' - Enable Admin Settings Menu
								If nAdmin <> PARENTS Then    
									    g_objIE.Document.getElementById("SET_KID_PASS").disabled = True 
									    g_objIE.Document.getElementById("SET_KID_PASS").style.color = "Grey"
									   	g_objIE.Document.getElementById("SIGN_CONFIG").disabled = True 
		                                g_objIE.Document.getElementById("SIGN_CONFIG").style.color = "Grey"		   
                                Else 
									   g_objIE.Document.getElementById("SET_KID_PASS").disabled = False 
									   g_objIE.Document.getElementById("SET_KID_PASS").style.color = HttpTextColor3
									   	g_objIE.Document.getElementById("SIGN_CONFIG").disabled = False 
		                                g_objIE.Document.getElementById("SIGN_CONFIG").style.color = HttpTextColor3   									   
							    End If
								Exit Do
							Case "Sign Out"
								strOwnerID = "NO_LOGON"
								nAdmin = "0"
                                Call WriteStrToFile(strDirectoryLCL & "\" & strFileSession, strOwnerID, 8, 1, nDebug)
								g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In"
							    g_objIE.Document.All("GreetingTitle").value = ""
								Exit Do
						End Select
                    Loop	
				Tsys0 = Date() & " " & Time()
		    Case "CHANGE_PASS"
			    Do
					' - Close Settings Panel
					For MenuX = 200 to 0 step - 4
						g_objIE.document.getElementById("divSettings").style.left = (MenuX - 200) & "px"
						'WScript.Sleep 1
					 Next
					nPressSettings = nPressSettings + 1
					' - Open New Password Window
					g_objIE.Document.All("ButtonHandler").Value = "None"
					If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" or strOwnerID = "NO_LOGON" Then 
						vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "Go to Sign in . . . "			    	: vvMsg(1,1) = "bold" 	: vvMsg(1,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 2)
					   Exit Do
					End If 
					Call IE_Hide(g_objIE)
					vvMsg(0,0) = "CHANGE PASSWORD FOR " & strOwnerID : vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor5
					If  Not IE_ChangePassword (g_objIE,vIE_Scale, vvMsg, 1, OldPasswd, NewPasswd,True,nDebug) Then Call IE_Unhide(g_objIE) :  Exit Do : End If
					OldPasswdMD5 = MD5Hash(Lcase(strOwnerID) & OldPasswd)
					
					strLine = Date() & " " & Time() & ": Login: " & strOwnerID & " CHANGE Password: " & OldPasswdMD5
					nAcctIndex = GetObjectLineNumber(vAccount,UBound(vAccount),strOwnerID, False) - 1
					If  nAcctIndex < 0 Then
						vvMsg(0,0) = "SOMETHING WENT WRONG"           :  vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "Signout and Sign In again"      :  vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Wrong Password", vvMsg, 2)
						Exit Do
					End If
					On Error Resume Next
					Err.Clear
					If OldPasswdMD5 <> Split(vAccount(nAcctIndex),",")(2) Then 
						Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", False)
						Call WriteStrToFile(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", strLine & "   FAILED", 100, 4, nDebug)
						Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", True)
						If Err.Number > 0 Then 
							MsgBox "Error: Line 781: " & Err.Description & chr(13) &_
													"strOwnerID=" & strOwnerID & chr(13) &_
													"AcctIndex=" & nAcctIndex & chr(13) &_
													"vAccount(nAcctIndex)=" & vAccount(nAcctIndex) & chr(13)
							Exit Do
						End If 
						vvMsg(0,0) = "WRONG PASSWORD!"      :  vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "TRY AGAIN . . ."      :  vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Wrong Password", vvMsg, 2)
						On Error Goto 0
						Exit Do
					End If
					' - Calculate ne Password Hash
					PasswdMD5 = MD5Hash(Lcase(strOwnerID) & NewPasswd)
					' - Create new Record for User strOwnerID
					strAccountRecord = strOwnerID & "," & Split(vAccount(nAcctIndex),",")(1) & "," & PasswdMD5
					For i = 3 to UBound(Split(vAccount(nAcctIndex),","))
					    strAccountRecord = strAccountRecord & "," & Split(vAccount(nAcctIndex),",")(i)
					Next 
					' - Write new record to strAccount File
'					MsgBox "OldPasswd: " & OldPasswd & chr(13) &_
'					       "NewPasswd: " & NewPasswd & chr(13) &_
'					       "User:      " & strOwnerID & chr(13) &_
'					       "OldHash:   " & OldPasswdMD5 & chr(13) &_
'					       "NewHash:   " & PasswdMD5
'				    MsgBox strAccountRecord
'					MsgBox strFileAccount
					vAccount(nAcctIndex) = strAccountRecord
					Call FileHidden(strFileAccount, False)
					Call FindAndReplaceStrInFile(strFileAccount, OldPasswdMD5, strAccountRecord, 1)
					Call FileHidden(strFileAccount, True)
					' - Add record to SignIn log lfile
					Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", False)
					Call WriteStrToFile(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", strLine & "   OK", 100, 4, nDebug)
					Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", True)
					On Error Goto 0
					Call IE_Unhide(g_objIE)
					Exit Do					
				Loop 
		    Case "SET_KID_PASS"
			    Do
					' - Close Settings Panel
					For MenuX = 200 to 0 step - 4
						g_objIE.document.getElementById("divSettings").style.left = (MenuX - 200) & "px"
						'WScript.Sleep 1
					 Next
					nPressSettings = nPressSettings + 1
					' - Open New Password Window
					g_objIE.Document.All("ButtonHandler").Value = "None"
					If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" or strOwnerID = "NO_LOGON" Then 
						vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "Go to Sign in . . . "			    	: vvMsg(1,1) = "bold" 	: vvMsg(1,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 2)
					   Exit Do
					End If 
					Call IE_Hide(g_objIE)
    				' - Choose Kids account you want to manage
					If IE_PromptForKidAccount(vIE_Scale, vDir, strOwnerID, strAccountID, vAccount, nDebug) = 0 Then 
					    Call IE_Unhide(g_objIE)											  
						Exit Do 
					End If
					vvMsg(0,0) = "CHANGE YOUR PASSWORD FOR " & strAccountID : vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor5
					If  Not IE_ChangePassword (g_objIE,vIE_Scale, vvMsg, 1, OldPasswd, NewPasswd,True,nDebug) Then Call IE_Unhide(g_objIE) :  Exit Do : End If
					OldPasswdMD5 = MD5Hash(Lcase(strOwnerID) & OldPasswd)
					
					strLine = Date() & " " & Time() & ": Admin: " & strOwnerID & "CHANGE Password for account: " & strAccountID
					nOwnerIndex = GetObjectLineNumber(vAccount,UBound(vAccount),strOwnerID, False) - 1
					nAcctIndex = GetObjectLineNumber(vAccount,UBound(vAccount),strAccountID, False) - 1
					
					If  nAcctIndex < 0 or nOwnerIndex < 0 Then
						vvMsg(0,0) = "SOMETHING WENT WRONG"           :  vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "Signout and Sign In again"      :  vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Wrong Password", vvMsg, 2)
						Exit Do
					End If
					On Error Resume Next
					Err.Clear
					If OldPasswdMD5 <> Split(vAccount(nOwnerIndex),",")(2) Then 
						Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", False)
						Call WriteStrToFile(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", strLine & "   FAILED", 100, 4, nDebug)
						Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", True)
						If Err.Number > 0 Then 
							MsgBox "Error: Line 861: " & Err.Description & chr(13) &_
													"strOwnerID=" & strOwnerID & chr(13) &_
													"nOwnerIndex=" & nOwnerIndex & chr(13) &_
													"vAccount(nOwnerIndex)=" & vAccount(nOwnerIndex) & chr(13)
							Exit Do
						End If 
						vvMsg(0,0) = "WRONG PASSWORD!"      :  vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "TRY AGAIN . . ."      :  vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2
						Call IE_MSG(g_objIE,vIE_Scale, "Wrong Password", vvMsg, 2)
						On Error Goto 0
						Exit Do
					End If
					' - Calculate ne Password Hash
					PasswdMD5 = MD5Hash(Lcase(strAccountID) & NewPasswd)
					' - Create new Record for User strOwnerID
					strAccountRecord = strAccountID & "," & Split(vAccount(nAcctIndex),",")(1) & "," & PasswdMD5
					For i = 3 to UBound(Split(vAccount(nAcctIndex),","))
					    strAccountRecord = strAccountRecord & "," & Split(vAccount(nAcctIndex),",")(i)
					Next 
					' - Write new record to strAccount File
'					MsgBox "OldPasswd: " & OldPasswd & chr(13) &_
'					       "NewPasswd: " & NewPasswd & chr(13) &_
'					       "User:      " & strAccountID & chr(13) &_
'					       "OldHash:   " & OldPasswdMD5 & chr(13) &_
'					       "NewHash:   " & PasswdMD5
'				    MsgBox strAccountRecord
    				vAccount(nAcctIndex) = strAccountRecord
					Call FileHidden(strFileAccount, False)
					Call FindAndReplaceStrInFile(strFileAccount, strAccountID & ",", strAccountRecord, 1)
					Call FileHidden(strFileAccount, True)
					' - Add record to SignIn log lfile
					Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", False)
					Call WriteStrToFile(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", strLine & "   OK", 100, 4, nDebug)
					Call FileHidden(strSrvDirectory & "\Data\" & "SignIn-" & strHost & ".log", True)
					On Error Goto 0
					Call IE_Unhide(g_objIE)
					Exit Do					
				Loop 				
			Case "SETTINGS"
				g_objIE.Document.All("ButtonHandler").Value = "None"
			    Select Case nPressSettings Mod 2
				    Case 0
							 For MenuX = 0 to 200 step 2
							    g_objIE.document.getElementById("divSettings").style.left = (MenuX - 200) & "px"
								'WScript.Sleep 1
							 Next
							 
							 nPressSettings = nPressSettings + 1
					Case Else
							 For MenuX = 200 to 0 step - 2
							    g_objIE.document.getElementById("divSettings").style.left = (MenuX - 200) & "px"
								'WScript.Sleep 1
							 Next
						nPressSettings = nPressSettings + 1
				End Select				
			Case "TIMER"
				g_objShell.SendKeys "% "
				wscript.sleep IE_PAUSE  
				g_objShell.SendKeys "n"
				wscript.sleep IE_PAUSE  
    			Call LaunchScript("TIMER",vInventory, nInventory,"", False)
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
			Case "TODO_LIST"
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
   				If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" Then 
					vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "TO EDIT YOUR TO-DO LIST"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
					vvMsg(2,0) = "Sign in . . . "			    		: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
 					Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 3)
				Else
					Call IE_Hide(g_objIE)
					Call LaunchScript("CHALLENGE",vInventory, nInventory,strPID & " " & strOwnerID, True)
					Call IE_Unhide(g_objIE)
					' g_objShell.AppActivate strPID					
				End If 
			Case "START"
				If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" Then 
					vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "TO START NEW SESSION"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
					vvMsg(2,0) = "Sign in . . . "			    		: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
 					Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 3)
				Else
					g_objShell.SendKeys "% "
					wscript.sleep IE_PAUSE  
					g_objShell.SendKeys "n"
					wscript.sleep IE_PAUSE
					Call LaunchScript("TASK_SET_SCRUSR",vInventory, nInventory,strWinUser, True)
					Call LaunchScript("START",vInventory, nInventory,strPID, False)
				End If
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
			Case "STOP"
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
    			If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" Then 
					vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "TO STOP ANY SESSION"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
					vvMsg(2,0) = "Sign in . . . "			    		: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
 					Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 3)
				Else
					Call IE_Hide(g_objIE)  
					Call LaunchScript("STOP",vInventory, nInventory,strPID, True)
					Call IE_Unhide(g_objIE)	
					' g_objShell.AppActivate strPID									
				End If
			Case "ALL"
				Call IE_Hide(g_objIE) 
				Call LaunchScript("DAY_STAT",vInventory, nInventory,strPID, True)
				Call IE_Unhide(g_objIE)
				' g_objShell.AppActivate strPID					
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
			Case "ACTIVE"
				Call IE_Hide(g_objIE)
				Call LaunchScript("ACTIVE",vInventory, nInventory,strPID, True)
				Call IE_Unhide(g_objIE)
                ' g_objShell.AppActivate strPID									
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
			Case "SERVER"
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()	
                Do	
					If Not objFSO.FileExists(strFileUnattend) Then 
						vvMsg(0,0) = "LAUNCHER WILL TRY TO RECONNECT TO THE SERVER"	: vvMsg(0,1) = "bold" : vvMsg(0,2) = HttpTextColor1
						vvMsg(1,0) = "Continue?"			                        : vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2					
						If Not IE_CONT(g_objIE,vIE_Scale, "Reconnecting to the Server",vvMsg,2,0) Then Exit Do 
                    Else 
					    objFSO.DeleteFile strFileUnattend, True
                    End If 
                    Call WriteStrToFile(strRestartFile, "START:", 1, 1,0)					
					strProcessName = "starting_my_vm_" & My_Random(1,999999)
					Call WriteProgressToFile(strProcessName, 5)
					Call AdminLaunchScript("RESTARTSERVICE",vInventory, nInventory, "ClientReportSvc" & " " & strProcessName)
					Call IE_Hide(g_objIE)
					Call LaunchScript("PROGRESS_REPORT",vInventory, nInventory, strProcessName & " ""Trying to Connect to Server""" & " ""CONNECTING TO MEDIA SERVER...""" & " 20" , True)
                    IE_LAUNCHER = 1
					g_objIE.Document.All("ButtonHandler").Value = "CANCEL"
	                Exit Do
				Loop
			Case "BOOKS"
				If g_objIE.Document.getElementById("LOGIN").innerHTML = "Sign In" Then 
					vvMsg(0,0) = "YOU MUST SIGN IN FIRST" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
					vvMsg(1,0) = "TO EDIT YOUR TASKS"					: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
					vvMsg(2,0) = "Sign in . . . "			    		: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2
 					Call IE_MSG(g_objIE,vIE_Scale, "Unknown user", vvMsg, 3)
				Else
				    Call IE_Hide(g_objIE)
					Call LaunchScript("BOOKS",vInventory, nInventory,strPID & " " & strOwnerID , True)
                    Call IE_Unhide(g_objIE)
                    ' g_objShell.AppActivate strPID										
				End If
				g_objIE.Document.All("ButtonHandler").Value = "None"
				Tsys0 = Date() & " " & Time()
			Case "MNG_KIDS_TASK"
					Call LaunchScript("BOOKS",vInventory, nInventory,strPID & " " & strAccountID , True)
                    Call IE_Unhide(g_objIE)
                    ' g_objShell.AppActivate strPID										
					g_objIE.Document.All("ButtonHandler").Value = "None"
					Tsys0 = Date() & " " & Time()
			Case "MNG_KIDS_TODO"
					Call LaunchScript("CHALLENGE",vInventory, nInventory,strPID & " " & strAccountID, True)
                    Call IE_Unhide(g_objIE)
                    ' g_objShell.AppActivate strPID										
					g_objIE.Document.All("ButtonHandler").Value = "None"
					Tsys0 = Date() & " " & Time()
			Case "MNG_KIDS_STATUS"
					Call LaunchScript("PARENT",vInventory, nInventory,strPID & " " & strAccountID, True)
                    Call IE_Unhide(g_objIE)
                    ' g_objShell.AppActivate strPID					
					g_objIE.Document.All("ButtonHandler").Value = "None"
					Tsys0 = Date() & " " & Time()
			Case "UPDATE"
				g_objIE.Document.All("ButtonHandler").Value = "None"
			    Do
					If strServerStatus <> SERVER_UP Then Exit Do
					If vUpdate( nUpdate - 1 ) = strVersion Then 
						strUpdate = vUpdate( nUpdate - 1 )
						vvMsg(0,0) = "YOU HAVE THE LASTEST VERSION " & strVersion 				    	: vvMsg(0,1) = "normal" : vvMsg(0,2) = "Green"
						vvMsg(1,0) = "INSTALLED ON YOUR COMPUTER" 					    				: vvMsg(1,1) = "normal" : vvMsg(1,2) = "Green"
						vvMsg(2,0) = "Proceed with updating it anyway?"									: vvMsg(2,1) = "bold" : vvMsg(2,2) = HttpTextColor1			
 						If Not IE_CONT(g_objIE,vIE_Scale, "KidsShield Updater v1.0",vvMsg,3,0) Then Exit Do
					End If
					Call IE_Hide(g_objIE)
					Call LaunchScript("UPDATE",vInventory, nInventory,"", True)
					If objFSO.FileExists(strFileUnattend) Then 
					    nConnect = GetFileLineCountSelect(strFileUnattend, vUnAtt,"#","NULL","NULL",0)
						Select Case Split(vUnAtt(0),":")(0)
						   Case "CONTINUE"
						        nInventory = GetFileLineCountByGroup(strDirectoryWork & "\" & strFileInventory, vInventory,"Server","Client","",0)
							    g_objIE.Document.All("ButtonHandler").Value = "UPDATE"
								Exit Do
						   Case "RECONNECT"
						        nInventory = GetFileLineCountByGroup(strDirectoryWork & "\" & strFileInventory, vInventory,"Server","Client","",0)
							    g_objIE.Document.All("ButtonHandler").Value = "SERVER"
								Exit Do
						    Case "END"
							    If objFSO.FileExists(strFileUnattend) Then 
								    objFSO.DeleteFile strFileUnattend, True
								End If
						End Select
					End If 
					g_objIE.Document.All("ButtonHandler").Value = "CANCEL"
					IE_LAUNCHER = 1
					Exit Do
				Loop
			Case "CANCEL"
				g_objIE.quit
				Exit Do
		End Select
		Wscript.Sleep 100  
		Loop
    Set g_objIE = Nothing
    Set g_objShell = Nothing
End Function
'------------------------------------------------------------------------------------------------------------
'   Function IE_PromptForAdminAction(vIE_Scale, strDir, strOwnerID, byRef strAccountID, byRef vAcct, nDebug)
'------------------------------------------------------------------------------------------------------------
Function IE_PromptForAdminAction(vIE_Scale, vDir, strOwnerID, byRef strAccountID, byRef vAcct, nDebug)
    Dim strDirectoryLCL, strSrvDirectory, strDirectoryWork, strDirectoryUpdate, strDirectoryVandyke
	Dim g_objIE, g_objShell
	Dim nAccount, nInd
	Dim nRatioX, nRatioY, nFontSize_10, nFontSize_12, nButtonX, nButtonY, nA, nB
    Dim intX
    Dim intY
	Dim nCount
	Dim strLogin
	Dim IE_Menu_Bar
	Dim  IE_Border
	Dim vApprovals, UserConfigFile, vAccount
	Dim IE_Window_Title
	IE_PromptForAdminAction = 0
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
    Set g_objShell = WScript.CreateObject("WScript.Shell")
	'-----------------------------------------------------------------
	'  FOLDER LIST
	'-----------------------------------------------------------------
	strDirectoryLCL = vDir(0)
	strSrvDirectory = vDir(1)
    strDirectoryWork = vDir(2)
	strDirectoryUpdate = vDir(3)
    strDirectoryVandyke = vDir(4)
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  g_objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	'------------------------------------------
	'	GET NUMBER OF TASKS LINES
	'------------------------------------------
	nAccount = UBound(vAcct)
	nLIne = 0 
	For nInd = 0 to nAccount
	     If Left(vAcct(nInd),1) <> "$" and Left(vAcct(nInd),1) <> "#" Then 
		    nLine = nLine + 1
		End If
	Next
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) 
	intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	intXreal = vIE_Scale(0,2) 
	intYreal = vIE_Scale(1,2) 
	nRatioX = intX/1920
    nRatioY = intY/1080
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------	
	BackGroundColor = HttpBgColor1
	ButtonColorOff = HttpBgColor4
	ButtonColor = HttpBgColor2
	MainTextColor = HttpTextColor1	
	'----------------------------------------
	' IE EXPLORER OBJECTS
	'----------------------------------------
    Call Set_IE_obj (g_objIE)
	'----------------------------------------
	' MAIN VARIABLES OF THE GUI FORM
	'----------------------------------------
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	nBottom = Round(10 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellH = Round(35 * nRatioY,0)
	LoginTitleW = Round(400 * nRatioX,0)
	nLeft = Round(20 * nRatioX,0)
	CellW = LoginTitleW
	LoginTitleH = Round(40 * nRatioY,0)
	WindowH = IE_Menu_Bar + 2 * LoginTitleH + cellH * (4 + nLine) + nButtonY + nBottom
	WindowW = IE_Border + LoginTitleW
	If WindowW < 300 then WindowW = 300 End If
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)	
	AcctButtonY = Round(30 * nRatioY,0)	
  If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &_
											" intX=" & intX &_
											" intY=" & intY &_
											" RatioX=" & nRatioX &_
											" RatioY=" & nRatioY &_
											" Window Width=" & WindowW &_
											" Window Hight=" & WindowH End If
	
    g_objIE.Offline = True
    g_objIE.navigate "about:blank"
   	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = MainTextColor
    g_objIE.height = WindowH
    g_objIE.width = WindowW  
    g_objIE.document.Title = "Manage Accounts"
	g_objIE.Top = (intYreal - g_objIE.height)/2
	g_objIE.Left = (intXreal - g_objIE.width)/2
	g_objIE.Visible = False
    '-----------------------------------------------------------------
	' Set the header and NewDevice Button of the HTML Body   		
	'-----------------------------------------------------------------
	strLine = _
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: " & HttpBgColor1 & "; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""" &_
			"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & 2 * Int(LoginTitleW/3) & """>" & _
				"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
				";font-weight: normal;font-style: italic;"">"
	If strOwnerID <> "NO_LOGON" Then 
		strLine = strLine &_
				"&nbsp;&nbsp;You are logged in as <span style=""font-weight: bold;"">" & strOwnerID & "</span></span></p>"
	Else 
		strLine = strLine & "Sign In </span></p>"
	End If 
	
	strLine = strLine &_
		"</td>" & _
		"</tr></tbody></table>"
	'-----------------------------------------------------------------
	' USERS TITLE
	'-----------------------------------------------------------------
		strLine = strLine &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; top: " & LoginTitleH + CellH & "px; left: 0px;" &_
		"border-collapse: collapse; border-style: none ; background-color: none'; width: " & CellW & "px;"">" & _
			"<tbody>" &_
				"<tr>" & _
					"<td style="" border-style: none;"" class=""oa2"" height=""" & cellH & """ width=""" & CellW & """>" & _
						"<p style=""text-align: center; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & HttpTextColor1 &_
						";font-weight: bold;font-style: normal;"">CHOOSE ACCOUT TO MANAGE</p>" &_
					"</td>" & _
				"</tr>"
	nLine = 1
	'-----------------------------------------------------------------
	' SET USER BUTTONS
	'-----------------------------------------------------------------
	nInd = 0
	Call TrDebug ("SELECT USER ACCOUNT: DRAW BUTTON TEMPLATES","", objDebug, MAX_LEN, 3, nDebug)
	Redim vAccount(Ubound(vAcct))
	Do While nInd < nAccount
	    vAccount(nInd) = Split(vAcct(nInd),",")(0)
	    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then 
		    Call TrDebug ("Account(" & nInd & "): " & vAccount(nInd),"", objDebug, MAX_LEN, 1, nDebug)
			strLine = strLine &_
					"<tr>" & _
						"<td style="" border-style: none;"" height=""" & cellH & """ align=""center"" class=""oa2"">" & _
							"<button id='AccountButton_1_" & nInd & "' name='AccountButton_1' AccessKey='" & nInd & "'" & _
							"style=""border-style: None;""" &_
							" onclick=document.all('ButtonHandler').value='ACCT_" & nInd & "';" &_
							" onmouseenter=document.all('ButtonHandler').value='MOUSEIN_" & nInd & "';" &_
							" onmouseleave=document.all('ButtonHandler').value='MOUSEOUT_" & nInd & "';" &_																
							"></button>" & _
						"</td>" &_
					"</tr>"
			nLine = nLine + 1
		End If
		nInd = nInd + 1
	Loop
	'-----------------------------------------------------------------
	' SET "PENDING APPROVAL" BUTTON
	'-----------------------------------------------------------------
	strLine = strLine &_
				"<tr>" & _
					"<td style="" border-style: none;"" align=""center"" class=""oa2"" height=""" & cellH & """>" & _
					"</td>" &_
				"</tr>" &_
				"<tr>" & _
					"<td style="" border-style: none;"" align=""center"" class=""oa2"" height=""" & cellH & """>" & _
						"<button id='PENDING' name='Pending' AccessKey='P'" & _
						"style=""border-style: None;""" &_
						" onclick=document.all('ButtonHandler').value='PENDING';></button>" & _
					"</td>" &_
				"</tr>"
	strLine = strLine & "</tbody></table>"
	'------------------------------------------------------
	'  ACTION  BUTTONS
	'------------------------------------------------------
    strLine = strLine &_
			"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: " & LoginTitleH & "px;" &_
			" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: " & HttpBgColor2 &_
			"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
				"<tbody>" & _
					"<tr>" &_
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" &_
							Int(LoginTitleW/4) - nLeft & ";height:" & nButtonY & "px;'" &_
							"id='Button1' name='APPROVE' AccessKey='A' onclick=document.all('ButtonHandler').value='MNGTASK';><u>A</u>pprovals</button>" & _
						"</td>" & _
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _						
							"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" &_
							Int(LoginTitleW/4) - nLeft & ";height:" & nButtonY & "px;'" &_
							"id='Button2' name='MNG' AccessKey='T' onclick=document.all('ButtonHandler').value='MNG';><u>M</u>anage</button>" & _
						"</td>" & _						
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
							"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" &_
							Int(LoginTitleW/4) - nLeft & ";height:" & nButtonY & "px;'" &_
							"id='Button3' name='MNGTODO' AccessKey='T' onclick=document.all('ButtonHandler').value='MNGTODO';><u>T</u>o-Do List</button>" & _
						"</td>" & _						
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _						
							"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" &_
							Int(LoginTitleW/4) - nLeft & ";height:" & nButtonY & "px;'" &_
							"id='Button4' name='EXIT' AccessKey='E' onclick=document.all('ButtonHandler').value='Exit';><u>E</u>xit</button>" & _
						"</td>" & _
					"</tr>" &_ 
				"</tbody></table>"	
	strLine = strLine &_				
				"<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"				
	'------------------------------------------------------
	'   BOTTOM INFO BAR 
	'------------------------------------------------------
	strLine = strLine &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: " & HttpBgColor6 &_
		"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW & """>" & _
						"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
						"; font-weight: normal;font-style: italic;"">" & "&nbsp;&nbsp;" &_
						"<span style=""font-weight: bold;"">Today: </span> " & GetMyDate() & ", " & Time() & " " &_
						"</span></p>" &_
					"</td>" & _
		"</tr></tbody></table>"
	'-----------------------------------------------------------------
	' HTML Form Parameaters
	'-----------------------------------------------------------------
    g_objIE.Document.Body.innerHTML = strLine
    g_objIE.MenuBar = False
    g_objIE.StatusBar = False
    g_objIE.AddressBar = False
    g_objIE.Toolbar = False
    '----------------------------------------------------------------
    ' SET DEFAULT VALUES
	'----------------------------------------------------------------
	strPendingAccount = "None"
	nInd = 0
	Do While nInd < nAccount
		UserConfigFile = strSrvDirectory & "\Data\config-" & vAccount(nInd) & ".ini"
		Call GetFileLineCountByGroup(UserConfigFile, vApprovals, "Approvals","","",nDebug)
		nIndex = GetObjectLineNumber( vApprovals, Ubound(vApprovals), "Request ToDo Approval", True)
		If nIndex > 0 Then
			If Split(vApprovals(nIndex - 1),"=")(1) = "1," & Date() Then 
				Type_of_approval = "TO-DO"
				strPendingAccount = vAccount(nInd)
				Exit Do
			End If
		End If 
		nIndex = GetObjectLineNumber( vApprovals, Ubound(vApprovals), "Request Task Approval", True)
		If nIndex > 0 Then
			If Split(vApprovals(nIndex - 1),"=")(1) = "1" Then 
				Type_of_approval = "TASK"
				strPendingAccount = vAccount(nInd)
				Exit Do
			End If
		End If 
		nInd = nInd + 1
	Loop
    '----------------------------------------------------------------
    ' DRAW PENDING APPROVAL BUTTON
	'----------------------------------------------------------------
	g_objIE.document.getElementById("PENDING").style.borderRadius = "25px"
	g_objIE.document.getElementById("PENDING").style.FontFamily = "arial narrow"
	g_objIE.document.getElementById("PENDING").style.fontWeight = "bold"
	g_objIE.document.getElementById("PENDING").Style.Height = AcctButtonY	
    g_objIE.document.getElementById("PENDING").Style.Width = 2 * nButtonX	
    g_objIE.Document.getElementById("PENDING").innerHTML = "PENDING APPROVAL"
	If strPendingAccount = "None" Then 
		g_objIE.Document.getElementById("PENDING").style.backgroundcolor = HttpBgColor1
		g_objIE.Document.getElementById("PENDING").style.color = "Grey"
		g_objIE.Document.getElementById("PENDING").disabled = True
		strPendingAccount = vAccount(0)
	Else
		g_objIE.Document.getElementById("PENDING").style.backgroundcolor = "Green"
		g_objIE.Document.getElementById("PENDING").style.color = HttpTextColor1
	End If
    '----------------------------------------------------------------
    ' DRAW USERS BUTTINS
	'----------------------------------------------------------------	
	Call TrDebug ("SELECT USER ACCOUNT: SET BUTTONS STYLE","", objDebug, MAX_LEN, 3, nDebug)
    For nInd = 0 to nAccount - 1
	    Call TrDebug ("POINT1: Account(" & nInd & "): " & vAccount(nInd),"", objDebug, MAX_LEN, 1, nDebug)
    	If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then 
		    Call TrDebug ("POINT2: Account(" & nInd & "): " & vAccount(nInd),"", objDebug, MAX_LEN, 1, nDebug)
			g_objIE.Document.getElementById("AccountButton_1_" & nInd ).innerHTML = UCase(vAccount(nInd))
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.borderRadius = "25px"
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.Height = AcctButtonY
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.fontFamily = "arial narrow"
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.fontSize = 14
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.Width = 2 * nButtonX				
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.fontWeight = "bold"
			If vAccount(nInd) = strPendingAccount Then 
				g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColor		
				g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = HttpTextColor2		
				g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked"
			Else 
	   		    g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
    			g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = BackGroundColor		
    		End If
	    End If 
	Next
	g_objIE.Document.getElementById("Button1").innerHTML = "APPROVALS"
	g_objIE.Document.getElementById("Button2").innerHTML = "MANAGE"
	g_objIE.Document.getElementById("Button3").innerHTML = "TO-DO LIST"
	g_objIE.Document.getElementById("Button4").innerHTML = "EXIT"
    g_objIE.document.getElementById("Button1").Style.fontFamily = "arial narrow"
	g_objIE.document.getElementById("Button1").Style.fontSize = 12
    g_objIE.document.getElementById("Button2").Style.fontFamily = "arial narrow"
	g_objIE.document.getElementById("Button2").Style.fontSize = 12
    g_objIE.document.getElementById("Button3").Style.fontFamily = "arial narrow"
	g_objIE.document.getElementById("Button3").Style.fontSize = 12
    g_objIE.document.getElementById("Button4").Style.fontFamily = "arial narrow"
	g_objIE.document.getElementById("Button4").Style.fontSize = 12
	
    Do
        WScript.Sleep 100
    Loop While g_objIE.Busy
    Call IE_Unhide(g_objIE)
    '-------------------------------------------------------------------
    ' Capture Result From HTML Form
	'-------------------------------------------------------------------
    Do
        On Error Resume Next
			If g_objIE.width <> WindowW Then g_objIE.width = WindowW End If
			If g_objIE.height <> WindowH Then g_objIE.height = WindowH End If
			Err.Clear
            szNothing = g_objIE.Document.All("ButtonHandler").Value
            if Err.Number <> 0 then exit function
        Select Case Split(szNothing,"_")(0)
		    Case "MOUSEIN"
			    g_objIE.Document.All("ButtonHandler").Value = "Nothing"
			    nButton = Split(szNothing,"_")(1)
				nInd = 0
				Do While nInd < nAccount
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then				
						Do
							If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then Exit Do
							If nInd = Int(nButton) Then g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColor : Exit Do : End If
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
							Exit Do
						Loop
					End If
					nInd = nInd + 1
				Loop
			Case "MOUSEOUT"
			    g_objIE.Document.All("ButtonHandler").Value = "Nothing"
			    nButton = Split(szNothing,"_")(1)
				nInd = 0
				Do While nInd < nAccount
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						Do
							If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then Exit Do
							If nInd = Int(nButton) Then g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff : Exit Do : End If
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
							Exit Do
						Loop
					End If
				    nInd = nInd + 1					
				Loop
            Case "ACCT"
				g_objIE.Document.All("ButtonHandler").Value = "Nothing"
			    nButton = Split(szNothing,"_")(1)
				nInd = 0
				Do While nInd < nAccount
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						If nInd = Int(nButton) Then
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColor		
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = HttpTextColor2		
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked"
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.fontWeight = "bold"
						Else 
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "unchecked"						
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = BackGroundColor		
						End If 
					End If 
					nInd = nInd + 1					
				Loop	
            Case "PENDING"
				strAccountID = strPendingAccount
				Select Case Type_of_approval
					Case "TASK"
						IE_PromptForAdminAction = 1
					Case "TO-DO"
						IE_PromptForAdminAction = 2
				End Select
                g_objIE.Quit
				Set g_objIE = Nothing
                Set g_objShell = Nothing
                Exit function
			Case "Exit"
                IE_PromptForAdminAction = 0
                g_objIE.Quit
                Set g_objIE = Nothing
				Set g_objShell = Nothing
                Exit function
            Case "MNGTASK"
                '-------------------------------------------------------------------
                ' Capture Selected Account from the List
	            '-------------------------------------------------------------------
				For nInd = 0 to nAccount - 1
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then 
							strAccountID = vAccount(nInd)
							Exit For
						End If 
					End If
				Next
                IE_PromptForAdminAction = 1
                g_objIE.Quit
				Set g_objIE = Nothing
                Set g_objShell = Nothing
                Exit function
            Case "MNGTODO"
                '-------------------------------------------------------------------
                ' Capture Selected Account from the List
	            '-------------------------------------------------------------------
				For nInd = 0 to nAccount - 1
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then 
							strAccountID = vAccount(nInd)
							Exit For
						End If 
					End If
				Next
                IE_PromptForAdminAction = 2
                g_objIE.Quit
				Set g_objIE = Nothing
                Set g_objShell = Nothing
                Exit function
			Case "MNG"
                '-------------------------------------------------------------------
                ' Capture Selected Account from the List
	            '-------------------------------------------------------------------
				For nInd = 0 to nAccount - 1
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then 
							strAccountID = vAccount(nInd)
							Exit For
						End If 
					End If
				Next
                IE_PromptForAdminAction = 3
                g_objIE.Quit
				Set g_objIE = Nothing
                Set g_objShell = Nothing
                Exit function
		End Select
        On Error Goto 0
		WScript.Sleep 50
    Loop
End Function
'------------------------------------------------------------------------------------------------------------
'   Function IE_PromptForKidAccount(vIE_Scale, strDir, strOwnerID, byRef strAccountID, byRef vAcct, nDebug)
'------------------------------------------------------------------------------------------------------------
Function IE_PromptForKidAccount(vIE_Scale, vDir, strOwnerID, byRef strAccountID, byRef vAcct, nDebug)
    Dim strDirectoryLCL, strSrvDirectory, strDirectoryWork, strDirectoryUpdate, strDirectoryVandyke
	Dim g_objIE, g_objShell
	Dim nAccount, nInd
	Dim nRatioX, nRatioY, nFontSize_10, nFontSize_12, nButtonX, nButtonY, nA, nB
    Dim intX
    Dim intY
	Dim nCount
	Dim strLogin
	Dim IE_Menu_Bar
	Dim  IE_Border
	Dim vApprovals, UserConfigFile, vAccount
	Dim IE_Window_Title
	IE_PromptForKidAccount = 0
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
    Set g_objShell = WScript.CreateObject("WScript.Shell")
	'-----------------------------------------------------------------
	'  FOLDER LIST
	'-----------------------------------------------------------------
	strDirectoryLCL = vDir(0)
	strSrvDirectory = vDir(1)
    strDirectoryWork = vDir(2)
	strDirectoryUpdate = vDir(3)
    strDirectoryVandyke = vDir(4)
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  g_objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	'------------------------------------------
	'	GET NUMBER OF TASKS LINES
	'------------------------------------------
	nAccount = UBound(vAcct)
	nLIne = 0 
	For nInd = 0 to nAccount
	     If Left(vAcct(nInd),1) <> "$" and Left(vAcct(nInd),1) <> "#" Then 
		    nLine = nLine + 1
		End If
	Next
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,0) : IE_Border = vIE_Scale(0,1) 
	intY = vIE_Scale(1,0) : IE_Menu_Bar = vIE_Scale(1,1)
	intXreal = vIE_Scale(0,2) 
	intYreal = vIE_Scale(1,2) 
	nRatioX = intX/1920
    nRatioY = intY/1080
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------	
	BackGroundColor = HttpBgColor1
	ButtonColorOff = HttpBgColor4
	ButtonColor = HttpBgColor2
	MainTextColor = HttpTextColor1	
	'----------------------------------------
	' IE EXPLORER OBJECTS
	'----------------------------------------
    Call Set_IE_obj (g_objIE)
	'----------------------------------------
	' MAIN VARIABLES OF THE GUI FORM
	'----------------------------------------
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	nBottom = Round(10 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellH = Round(35 * nRatioY,0)
	LoginTitleW = Round(400 * nRatioX,0)
	nLeft = Round(20 * nRatioX,0)
	CellW = LoginTitleW
	LoginTitleH = Round(40 * nRatioY,0)
	WindowH = IE_Menu_Bar + 2 * LoginTitleH + cellH * (4 + nLine) + nButtonY + nBottom
	WindowW = IE_Border + LoginTitleW
	If WindowW < 300 then WindowW = 300 End If
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)	
	AcctButtonY = Round(30 * nRatioY,0)	
  If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &_
											" intX=" & intX &_
											" intY=" & intY &_
											" RatioX=" & nRatioX &_
											" RatioY=" & nRatioY &_
											" Window Width=" & WindowW &_
											" Window Hight=" & WindowH End If
	
    g_objIE.Offline = True
    g_objIE.navigate "about:blank"
   	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = MainTextColor
    g_objIE.height = WindowH
    g_objIE.width = WindowW  
    g_objIE.document.Title = "Manage Accounts"
	g_objIE.Top = (intYreal - g_objIE.height)/2
	g_objIE.Left = (intXreal - g_objIE.width)/2
	g_objIE.Visible = False
    '-----------------------------------------------------------------
	' Set the header and NewDevice Button of the HTML Body   		
	'-----------------------------------------------------------------
	strLine = _
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 &_
		"; background-color: " & HttpBgColor1 & "; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
		"<tbody>" & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""" &_
			"valign=""middle"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & 2 * Int(LoginTitleW/3) & """>" & _
				"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_				
				";font-weight: normal;font-style: italic;"">"
	If strOwnerID <> "NO_LOGON" Then 
		strLine = strLine &_
				"&nbsp;&nbsp;You are logged in as <span style=""font-weight: bold;"">" & strOwnerID & "</span></span></p>"
	Else 
		strLine = strLine & "Sign In </span></p>"
	End If 
	
	strLine = strLine &_
		"</td>" & _
		"</tr></tbody></table>"
	'-----------------------------------------------------------------
	' USERS TITLE
	'-----------------------------------------------------------------
		strLine = strLine &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; top: " & LoginTitleH + CellH & "px; left: 0px;" &_
		"border-collapse: collapse; border-style: none ; background-color: none'; width: " & CellW & "px;"">" & _
			"<tbody>" &_
				"<tr>" & _
					"<td style="" border-style: none;"" class=""oa2"" height=""" & cellH & """ width=""" & CellW & """>" & _
						"<p style=""text-align: center; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & HttpTextColor1 &_
						";font-weight: bold;font-style: normal;"">CHOOSE ACCOUT TO MANAGE</p>" &_
					"</td>" & _
				"</tr>"
	nLine = 1
	'-----------------------------------------------------------------
	' SET USER BUTTONS
	'-----------------------------------------------------------------
	nInd = 0
	Call TrDebug ("SELECT USER ACCOUNT: DRAW BUTTON TEMPLATES","", objDebug, MAX_LEN, 3, nDebug)
	Redim vAccount(Ubound(vAcct))
	Do While nInd < nAccount
	    vAccount(nInd) = Split(vAcct(nInd),",")(0)
	    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then 
		    Call TrDebug ("Account(" & nInd & "): " & vAccount(nInd),"", objDebug, MAX_LEN, 1, nDebug)
			strLine = strLine &_
					"<tr>" & _
						"<td style="" border-style: none;"" height=""" & cellH & """ align=""center"" class=""oa2"">" & _
							"<button id='AccountButton_1_" & nInd & "' name='AccountButton_1' AccessKey='" & nInd & "'" & _
							"style=""border-style: None;""" &_
							" onclick=document.all('ButtonHandler').value='ACCT_" & nInd & "';" &_
							" onmouseenter=document.all('ButtonHandler').value='MOUSEIN_" & nInd & "';" &_
							" onmouseleave=document.all('ButtonHandler').value='MOUSEOUT_" & nInd & "';" &_																
							"></button>" & _
						"</td>" &_
					"</tr>"
			nLine = nLine + 1
		End If
		nInd = nInd + 1
	Loop
	'------------------------------------------------------
	'  ACTION  BUTTONS
	'------------------------------------------------------
    strLine = strLine &_
			"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: " & LoginTitleH & "px;" &_
			" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: " & HttpBgColor2 &_
			"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
				"<tbody>" & _
					"<tr>" &_
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""center"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"</td>" & _
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _						
						"</td>" & _						
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _
						"</td>" & _						
						"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & Int(LoginTitleW/4) & """>" & _						
							"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" &_
							Int(LoginTitleW/4) - nLeft & ";height:" & nButtonY & "px;'" &_
							"id='Button4' name='EXIT' AccessKey='E' onclick=document.all('ButtonHandler').value='Exit';><u>E</u>xit</button>" & _
						"</td>" & _
					"</tr>" &_ 
				"</tbody></table>"	
	strLine = strLine &_				
				"<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"				
	'------------------------------------------------------
	'   BOTTOM INFO BAR 
	'------------------------------------------------------
	strLine = strLine &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; bottom: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: " & HttpBgColor6 &_
		"; height: " & LoginTitleH & "px; width: " & LoginTitleW & "px;"">" & _
			"<tbody>" & _
				"<tr>" &_
					"<td style=""border-style: none; background-color: " & HttpBgColor5 & ";""align=""right"" class=""oa1"" height=""" & LoginTitleH & """ width=""" & LoginTitleW & """>" & _
						"<p><span style="" font-size: " & nFontSize_10 & ".0pt; font-family: 'Helvetica'; color: " & HttpTextColor3 &_
						"; font-weight: normal;font-style: italic;"">" & "&nbsp;&nbsp;" &_
						"<span style=""font-weight: bold;"">Today: </span> " & GetMyDate() & ", " & Time() & " " &_
						"</span></p>" &_
					"</td>" & _
		"</tr></tbody></table>"
	'-----------------------------------------------------------------
	' HTML Form Parameaters
	'-----------------------------------------------------------------
    g_objIE.Document.Body.innerHTML = strLine
    g_objIE.MenuBar = False
    g_objIE.StatusBar = False
    g_objIE.AddressBar = False
    g_objIE.Toolbar = False
    '----------------------------------------------------------------
    ' DRAW USERS BUTTINS
	'----------------------------------------------------------------	
	Call TrDebug ("SELECT USER ACCOUNT: SET BUTTONS STYLE","", objDebug, MAX_LEN, 3, nDebug)
    For nInd = 0 to nAccount - 1
	    Call TrDebug ("POINT1: Account(" & nInd & "): " & vAccount(nInd),"", objDebug, MAX_LEN, 1, nDebug)
    	If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then 
		    Call TrDebug ("POINT2: Account(" & nInd & "): " & vAccount(nInd),"", objDebug, MAX_LEN, 1, nDebug)
			g_objIE.Document.getElementById("AccountButton_1_" & nInd ).innerHTML = UCase(vAccount(nInd))
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.borderRadius = "25px"
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.Height = AcctButtonY
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.fontFamily = "arial narrow"
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.fontSize = 14
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).Style.Width = 2 * nButtonX				
			g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.fontWeight = "bold"
			If vAccount(nInd) = strPendingAccount Then 
				g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColor		
				g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = HttpTextColor2		
				g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked"
			Else 
	   		    g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
    			g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = BackGroundColor		
    		End If
	    End If 
	Next
	g_objIE.Document.getElementById("Button4").innerHTML = "EXIT"
    g_objIE.document.getElementById("Button4").Style.fontFamily = "arial narrow"
	g_objIE.document.getElementById("Button4").Style.fontSize = 12
	
    Do
        WScript.Sleep 100
    Loop While g_objIE.Busy
    Call IE_Unhide(g_objIE)
    '-------------------------------------------------------------------
    ' Capture Result From HTML Form
	'-------------------------------------------------------------------
    Do
        On Error Resume Next
			If g_objIE.width <> WindowW Then g_objIE.width = WindowW End If
			If g_objIE.height <> WindowH Then g_objIE.height = WindowH End If
			Err.Clear
            szNothing = g_objIE.Document.All("ButtonHandler").Value
            if Err.Number <> 0 then exit function
        Select Case Split(szNothing,"_")(0)
		    Case "MOUSEIN"
			    g_objIE.Document.All("ButtonHandler").Value = "Nothing"
			    nButton = Split(szNothing,"_")(1)
				nInd = 0
				Do While nInd < nAccount
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then				
						Do
							If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then Exit Do
							If nInd = Int(nButton) Then g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColor : Exit Do : End If
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
							Exit Do
						Loop
					End If
					nInd = nInd + 1
				Loop
			Case "MOUSEOUT"
			    g_objIE.Document.All("ButtonHandler").Value = "Nothing"
			    nButton = Split(szNothing,"_")(1)
				nInd = 0
				Do While nInd < nAccount
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						Do
							If g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked" Then Exit Do
							If nInd = Int(nButton) Then g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff : Exit Do : End If
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
							Exit Do
						Loop
					End If
				    nInd = nInd + 1					
				Loop
            Case "ACCT"
				g_objIE.Document.All("ButtonHandler").Value = "Nothing"
			    nButton = Split(szNothing,"_")(1)
				nInd = 0
				Do While nInd < nAccount
				    If Left(vAccount(nInd),1) <> "$" and Left(vAccount(nInd),1) <> "#" Then
						If nInd = Int(nButton) Then
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColor		
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = HttpTextColor2		
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "checked"
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.fontWeight = "bold"
							strAccountID = vAccount(nInd)
							IE_PromptForKidAccount = 1
							g_objIE.Quit
							Set g_objIE = Nothing
							Set g_objShell = Nothing
							Exit function
						Else 
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.backgroundcolor = ButtonColorOff
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).Value = "unchecked"						
							g_objIE.document.getElementById("AccountButton_1_" & nInd ).style.color = BackGroundColor		
						End If 
					End If 
					nInd = nInd + 1					
				Loop	
			Case "Exit"
                IE_PromptForKidAccount = 0
                g_objIE.Quit
                Set g_objIE = Nothing
				Set g_objShell = Nothing
                Exit function
		End Select
        On Error Goto 0
		WScript.Sleep 50
    Loop
End Function
'----------------------------------------------------------------------------------
'    Function SetScreenUser
'----------------------------------------------------------------------------------
Function SetScreenUser(strFile, byRef StrWinUser, nDebug)
	Dim vLine
	Dim strScreenUser, strUserProfile
	Dim nCount
	Dim objEnvar, objScreenUserFile
	SetScreenUser = False
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	If nDebug = 1 Then 
		objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  SetScreenUser: CURRENT WINDOWS USER: " & strScreenUser
	End If
	On Error resume next
	set objScreenUserFile = objFSO.OpenTextFile(strFile,ForWriting,True)
	If Err.Number <> 0 Then
		objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  SetScreenUser ERROR: CAN'T CREATE SCREENUSER_tmp.dat" & chr(13) &_
		                                             "-------------------------------> " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description 
		Err.Clear
	End If
	objScreenUserFile.WriteLine strScreenUser
	If Err.Number <> 0 Then
		objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  SetScreenUser ERROR: CAN'T WRITE TO SCREENUSER_tmp.dat" & chr(13) &_
		                                             "-------------------------------> " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description 
		Err.Clear
	End If
	On Error goto 0
	If Err.Number = 0 Then 	
		objScreenUserFile.Close 
	End If
	
	If strScreenUser = GetScreenUser(strFile, nDebug) Then 
		SetScreenUser = True 
		strWinUser = strScreenUser
	End If
	set objEnvar = Nothing
End Function
'------------------------------------------------------------------------------------------------------------------
' Function GetConnectStatus (ByVal strNetworkStatus, ByVal strServerStatus, nDebug) returns 
'------------------------------------------------------------------------------------------------------------------
Function GetConnectStatus (ByRef strNetworkStatus, ByRef strServerStatus, ByVal strServerIP, ByVal strGatewayIP, ByVal strFileConnect, nSleep, nRetry,bSave, nDebug)
Dim nInd
Dim objConnect
	strServerStatus = OFF_LINE
	strNetworkStatus = OFF_LINE
		nInd = 0
	Do While strServerStatus <> ON_LINE
		strServerPing = GetCmdPing(strServerIP, nDebug)		
		strGatewayPing = GetCmdPing(strGatewayIP, nDebug)		
		If strServerPing = "....." and strGatewayPing = "....." Then 
			strNetworkStatus = OFF_LINE
			strServerStatus = UNKNOWN
		End If
		If strServerPing = "....." and strGatewayPing = "!!!!!" Then 
			strNetworkStatus = ON_LINE
			strServerStatus = OFF_LINE
		End If
		If strServerPing = "!!!!!" and strGatewayPing = "!!!!!" Then 
			strNetworkStatus = ON_LINE
			strServerStatus = ON_LINE
		End If
		If strServerPing = "!!!!!" and strGatewayPing = "....." Then 
			strNetworkStatus = NO_INTERNET
			strServerStatus = ON_LINE
		End If
		If bSave Then
			Do
				On Error Resume Next
				Set objConnect = objFSO.OpenTextFile(strFileConnect, ForWriting,True)
 				If Err.Number = 0 Then 
					objConnect.WriteLine "Network," & strNetworkStatus
					objConnect.WriteLine "Server," & strServerStatus
					objConnect.Close
					Err.Clear
					On Error goto 0
					Exit Do
				Else
					objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WARNING:  Can't WRITE TO connectivity.dat. FILE IS BUSY. RETRY in 1 sec" 
					Err.Clear
					On Error goto 0
					wscript.sleep 1000
				End If
			Loop
		End If
		nInd = nInd + 1
		If strServerStatus = OFF_LINE Then 
			objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WARNING:  CAN'T REACH SERVER"  
		End If
		If strNetworkStatus <> ON_LINE Then 
			objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WARNING:  NETWORK NOT ON_LINE. STATUS: " & strNetworkStatus 
		End If
		If nInd = nRetry Then Exit Do End if
		If strServerStatus = ON_LINE Then Exit Do End If
		wscript.sleep nSleep
	Loop	
End Function
'------------------------------------------------------------------------------------------------------------------
' Function GetServerStatus (ByVal strDummyServerFile, ByVal strServerStatus, nDebug) returns 
'------------------------------------------------------------------------------------------------------------------
Function GetServerStatus (ByVal strDummyServerFile, ByRef strNetworkStatus, ByRef strServerStatus, strFileConnect, nSleep, nRetry, bSave,nDebug)
Dim nInd
Dim objConnect
	nInd = 0
	Do While strServerStatus <> SERVER_UP
		If objFSO.FileExists(strDummyServerFile) Then 
			strServerStatus = SERVER_UP
			If bSave Then
				Do
					On Error Resume Next
					Set objConnect = objFSO.OpenTextFile(strFileConnect, ForWriting,True)
					If Err.Number = 0 Then 
						objConnect.WriteLine "Network," & strNetworkStatus
						objConnect.WriteLine "Server," & strServerStatus
						objConnect.Close
						Err.Clear
						On Error goto 0
						Exit Do
					Else
						objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WARNING:  Can't WRITE TO connectivity_tmp.dat. FILE IS BUSY. RETRY in 1 sec" 
						Err.Clear
						On Error goto 0
						wscript.sleep 1000
					End If
				Loop
			End If
			Exit Do
		Else
			objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": WARNING:  SERVER IS LOADING... FILESYSTEM IS NOT READY" 
		End If
		nInd = nInd + 1
		If nInd = nRetry Then 
			objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": ERROR:  SERVER FILESYSTEM IS NOT READY. CONTINUE IN OFF-LINE MODE"  
			Exit Do 	
		End if
		If strServerStatus = SERVER_UP Then Exit Do End If
		wscript.sleep nSleep
	Loop
End Function
'----------------------------------------------------------------------------------
'    Function WriteScreenUserToServer
'----------------------------------------------------------------------------------
Function WriteScreenUserToServer(strFile, byRef StrWinUser, nDebug)
Const ForAppending = 8
Const ForWriting = 2
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
	WriteScreenUserToServer = False
	On Error resume next
	set objScreenUserFile = objFSO.OpenTextFile(strFile,ForWriting,True)
	If Err.Number <> 0 Then
		g_objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  WriteScreenUserToServer ERROR: CAN'T CREATE SCREENUSER_tmp.dat" & chr(13) &_
		                                             "-------------------------------> " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description 
		Err.Clear
		On Error goto 0
		Exit Function
	End If
	objScreenUserFile.WriteLine strWinUser
	If Err.Number <> 0 Then
		g_objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  WriteScreenUserToServer ERROR: CAN'T WRITE TO SCREENUSER_tmp.dat" & chr(13) &_
		                                             "-------------------------------> " & Err.Number & " Srce: " & Err.Source & " Desc: " &  Err.Description 
		Err.Clear
		On Error goto 0
		objScreenUserFile.Close 
		Exit Function
	End If
	On Error goto 0
	If Err.Number = 0 Then 	
		objScreenUserFile.Close 
	End If
	WriteScreenUserToServer = True 
End Function
'----------------------------------------------------------------------------------
'    Function GetLogonUser
'----------------------------------------------------------------------------------
Function GetLogonUser(nDebug)
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar, objScreenUserFile
GetLogonUser = "NO_LOGON"
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	If nDebug = 1 Then 
		objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ":  GetLogonUser: CURRENT WINDOWS USER: " & strScreenUser
	End If
GetLogonUser = strScreenUser
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUser
'----------------------------------------------------------------------------------
Function GetScreenUser(strFile, nDebug)
Dim vFileLines
Dim strScreenUser
Dim nLine
		GetScreenUser = ""
		If objFSO.FileExists(strFile) Then 
			nLine = GetFileLineCountSelect(strFile, vFileLines,"#","NULL","NULL",0)
			strScreenUser = vFileLines(0)
			Call TrDebug ("Current Win User:", strScreenUser, objDebug, MAX_LEN, 1, nDebug)
		Else 
			Call TrDebug ("GetScreenUser", "ERROR:  screenuser_tmp.dat NOT FOUND", objDebug, MAX_LEN, 1, nDebug)
		End If		
		GetScreenUser = strScreenUser
End Function
'-------------------------------------------------------------------------
' Function GetAccountStatus(vAccount,nAcctIndex,nAdmin,strBlock,nDebug)
'-------------------------------------------------------------------------
Function GetAccountStatus(ByRef vAccount,nAcctIndex,ByRef nAdmin, ByRef nBlock,nDebug)
Dim nAuto, strUser, nVector
Const FATHER = "Father"
Dim UserVector(8)
    UserVector(0) = "000"
	UserVector(1) = "001"
	UserVector(2) = "010"
	UserVector(3) = "011"
	UserVector(4) = "100"
	UserVector(5) = "101"
	UserVector(6) = "110"	
	UserVector(7) = "111"		
	nAdmin = 0
	nBlock = 0
	nAuto = 0
	strUser = Split(vAccount(nAcctIndex),",")(0)
	nVector = 4
	Do
		If strUser = FATHER then 
			nVector = 2
			Exit Do
		End If
		If Left(strUser,1) = "$" Then 
			nVector = 0
			Exit Do		
		End If 
		For n = 0 to 7
			If Split(vAccount(nAcctIndex),",")(1) = MD5Hash(LCase(strUser) & UserVector(n)) Then 
			   nVector = n
			   Exit For
			End If
		Next
		Exit Do
	Loop
    nBlock = Mid(UserVector(nVector),1,1)
    nAdmin = Mid(UserVector(nVector),2,1)
	nAuto = Mid(UserVector(nVector),3,1)
	GetAccountStatus = True
End Function
'------------------------------------------------------------------------------
'      Function ASKS USER TO ENTER PASSWORD
'------------------------------------------------------------------------------
 Function IE_PromptLoginPassword (objParentWin, vIE_Scale, vLine, nLine, ByRef strUsername, ByRef strPassword, Confirm, nDebug )
    Dim strPID
	Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim g_objIE, g_objShell
	intX = 1920
	intY = 1080
	Dim IE_Menu_Bar
	Dim  IE_Border
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	strPassword = "DO NOT MATCH"
	IE_PromptLoginPassword = False	
	
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nHeader = Round (12 * nRatioY,0)
	LineH = Round (12 * nRatioY,0)
	nTab = 20
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellW = Round(330 * nRatioX,0)
	ColumnW1 = Round(150 * nRatioX,0)
	CellH = 2 * (nLine + 7) * LineH
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
	If Confirm Then 
	    CellH = CellH + 3 * 2 * LineH 
		nOrder = 1
    Else 
	    nOrder = 0
	End If
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------		
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	InputBGColor = HttpBgColor4
	MainTextColor = HttpTextColor1
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = BackGroundColor
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	'----------------------------------------------------------
	'    TITLE
	'----------------------------------------------------------
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: "& HttpBgColor2 & ";" &_
		"width: " & CellW & "px;"">" & _
		"<tbody>"	
	For nInd = 0 to nLine - 1
		 If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
		strHTMLBody = strHTMLBody &_
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & CellW & """>" & _
				"<p style=""text-align: center; font-family: 'arial narrow';font-size: " & nFontSize_12 & ".0pt; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" &_
			"</td>" &_
		"</tr>"
	Next
	strHTMLBody = strHTMLBody & "</tbody></table>"
	
	'----------------------------------------------------------
	'    MAIN FORM FOR ENTERING LOGON AND PASSWORD
	'----------------------------------------------------------
	TableW = CellW
	ColumnW_1 = 3 * Int(TableW/3)
	ColumnW_2 = TableW - ColumnW_1
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: " & (nLine + 1) * LineH * 2 & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: none;;" &_
		"width: " & TableW & "px;"">" & _
		"<tbody>"		
	'----------------------------------------------------	
	'  ROW 1
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">LOGIN NAME</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		"</td>" &_
	"</tr>"		
	'----------------------------------------------------	
	'  ROW 2
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """ >" & _
			"<input name=UserName style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=25 tabindex=1>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='EXIT' name='Cancel' AccessKey='C' tabindex=" & nOrder + 4 & " onclick=document.all('ButtonHandler').value='Cancel';>CANCEL</button>" & _		    
		"</td>" &_		
	"</tr>"
	'----------------------------------------------------	
	'  ROW 3 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  ROW 4
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">PASSWORD</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
		"</td>" &_
	"</tr>"			
	'----------------------------------------------------	
	'  ROW 5
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
			"<input id='PASSWD' name=Password style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=2 " & _
			"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='OK' name='OK' AccessKey='C' tabindex=" & nOrder + 3 & " onclick=document.all('ButtonHandler').value='OK';>SIGN IN</button>" & _		    
		"</td>" &_				
	"</tr>"
	'----------------------------------------------------	
	'  ROW 6 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  CONFIRM PASSWORD ROW
	'----------------------------------------------------
	If Confirm Then 
		'----------------------------------------------------	
		'  ROW 7
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
				"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
				"; font-weight: bold;"">CONFIRM PASSWORD</p>" &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"</td>" &_
		"</tr>"			
		'----------------------------------------------------	
		'  ROW 8
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
				"<input id='PASSWD2' name=Password2 style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
				"; border-radius: 10px " &_
				"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=3 " & _
				"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """>" & _
			"</td>" &_				
		"</tr>"
	End If
	strHTMLBody = strHTMLBody & "</tbody></table>"
    strHTMLBody = strHTMLBody &_
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = "Login and Password"
	g_objIE.document.getElementById("OK").style.borderRadius = "10px"
	g_objIE.document.getElementById("EXIT").style.borderRadius = "10px"
	g_objIE.document.getElementById("OK").style.backgroundcolor = ButtonColor
	g_objIE.document.getElementById("EXIT").style.backgroundcolor = ButtonColor
	If Confirm Then
	    g_objIE.Document.getElementById("OK").innerHTML = "OK"
	Else 
	   	g_objIE.Document.getElementById("OK").innerHTML = "SIGN IN"
	End If
	
	g_objIE.Visible = False
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy	
	Set g_objShell = WScript.CreateObject("WScript.Shell")
	Call IE_Unhide(g_objIE)
	Call IE_GetPID(strPID, g_objIE.document.Title & " - " & IE_Window_Title, nDebug)
	g_objShell.AppActivate strPID
	g_objIE.Document.All("UserName").Focus
	g_objIE.Document.All("UserName").Value = strUsername
'    g_objIE.Document.body.addeventlistener "keydown", GetRef("KeyLA"), false
	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case strNothing
			Case "Cancel"
				' The user clicked Cancel. Exit the loop
				IE_PromptLoginPassword = False				
				Exit Do
			Case "OK"
				' strUsername = g_objIE.Document.All("Username").Value
				Select Case Confirm
					Case True
						if g_objIE.Document.All("Password").Value = g_objIE.Document.All("Password2").Value  and _
						   InStr(g_objIE.Document.All("Password").Value," ") = 0 and _
						   g_objIE.Document.All("Password").Value <> "" Then 
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = g_objIE.Document.All("Password").Value
							IE_PromptLoginPassword = True
							Exit Do
						Else
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = "DO NOT MATCH"
							IE_PromptLoginPassword = True
							Exit Do
						End If 
					Case False
							strUsername = g_objIE.Document.All("UserName").Value
							strPassword = g_objIE.Document.All("Password").Value
							IE_PromptLoginPassword = True
							Exit Do
				End Select
		End Select
	    Wscript.Sleep 200
    Loop
	g_objIE.quit
	Wscript.Sleep 200
	Set g_objIE = Nothing
	Set g_objShell = Nothing
End Function
'------------------------------------------------------------------------------
'      Function ASKS USER TO CHANGE ITS PASSWORD
'------------------------------------------------------------------------------
 Function IE_ChangePassword (objParentWin, vIE_Scale, vLine, nLine, ByRef strPassword, ByRef strNewPass, Confirm, nDebug )
    Dim strPID
	Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim g_objIE, g_objShell
	intX = 1920
	intY = 1080
	Dim IE_Menu_Bar
	Dim  IE_Border
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	strPassword = "DO NOT MATCH"
	IE_ChangePassword = False	
	
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	' This loop is required to allow the IE object to finish loading...
	Do
		WScript.Sleep 200
	Loop While g_objIE.Busy
	nHeader = Round (12 * nRatioY,0)
	LineH = Round (12 * nRatioY,0)
	nTab = 20
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_14 = Round(14 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	CellW = Round(330 * nRatioX,0)
	ColumnW1 = Round(150 * nRatioX,0)
	CellH = 2 * (nLine + 7) * LineH
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
	If Confirm Then 
	    CellH = CellH + 3 * 2 * LineH 
		nOrder = 1
    Else 
	    nOrder = 0
	End If
	WindowW = IE_Border + CellW
	WindowH = IE_Menu_Bar + CellH
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------		
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	InputBGColor = HttpBgColor4
	MainTextColor = HttpTextColor1
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = BackGroundColor
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	'----------------------------------------------------------
	'    TITLE
	'----------------------------------------------------------
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: 0px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor2 & "; background-color: "& HttpBgColor2 & ";" &_
		"width: " & CellW & "px;"">" & _
		"<tbody>"	
	For nInd = 0 to nLine - 1
		 If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
		strHTMLBody = strHTMLBody &_
		"<tr>" &_
			"<td style=""border-style: none; background-color: " & HttpBgColor2 & ";"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & CellW & """>" & _
				"<p style=""text-align: center; font-family: 'arial narrow';font-size: " & nFontSize_12 & ".0pt; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" &_
			"</td>" &_
		"</tr>"
	Next
	strHTMLBody = strHTMLBody & "</tbody></table>"
	
	'----------------------------------------------------------
	'    MAIN FORM FOR ENTERING LOGON AND PASSWORD
	'----------------------------------------------------------
	TableW = CellW
	ColumnW_1 = 3 * Int(TableW/3)
	ColumnW_2 = TableW - ColumnW_1
	strHTMLBody = strHTMLBody &_
		"<table border=""1"" cellpadding=""1"" cellspacing=""1"" style="" position: absolute; left: 0px; top: " & (nLine + 1) * LineH * 2 & "px;" &_
		" border-collapse: collapse; border-style: none; border width: 1px; border-color: " & HttpBgColor5 & "; background-color: none;;" &_
		"width: " & TableW & "px;"">" & _
		"<tbody>"		
	'----------------------------------------------------	
	'  ROW 1
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">CURRENT PASSWORD</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		"</td>" &_
	"</tr>"		
	'----------------------------------------------------	
	'  ROW 2
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """ >" & _
			"<input id='OLDPASS' name=OldPass style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" type=password AccessKey=p size=20 maxlength=25 tabindex=1>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='EXIT' name='Cancel' AccessKey='C' tabindex=" & nOrder + 4 & " onclick=document.all('ButtonHandler').value='Cancel';>CANCEL</button>" & _		    
		"</td>" &_		
	"</tr>"
	'----------------------------------------------------	
	'  ROW 3 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  ROW 4
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
			"; font-weight: bold;"">NEW PASSWORD</p>" &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
		"</td>" &_
	"</tr>"			
	'----------------------------------------------------	
	'  ROW 5
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
			"<input id='PASSWD' name=Password style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
			"; border-radius: 10px " &_
			"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=2 " & _
			"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
		"</td>" &_
		"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """ width=""" & ColumnW_1 & """>" & _
		   "<button style=""font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
			"; width:" & nButtonX & ";height:" & 2 * LineH & "; font-family: 'arial narrow';""" & _
			"id='OK' name='OK' AccessKey='C' tabindex=" & nOrder + 3 & " onclick=document.all('ButtonHandler').value='OK';>SIGN IN</button>" & _		    
		"</td>" &_				
	"</tr>"
	'----------------------------------------------------	
	'  ROW 6 (EMPTY)
	'----------------------------------------------------
	strHTMLBody = strHTMLBody & _
	"<tr>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
		"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  LineH & """>" & _
		"</td>" &_
	"</tr>"
	'----------------------------------------------------	
	'  CONFIRM PASSWORD ROW
	'----------------------------------------------------
	If Confirm Then 
		'----------------------------------------------------	
		'  ROW 7
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
				"<p style=""position: relative; left: " & Int(nTab/2) & "px; bottom: -3px; font-size: " & nFontSize_12 & ".0pt; font-family: 'arial narrow'; color: " & MainTextColor &_
				"; font-weight: bold;"">CONFIRM PASSWORD</p>" &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;""class=""oa1"" height=""" &  2 * LineH & """ >" & _
			"</td>" &_
		"</tr>"			
		'----------------------------------------------------	
		'  ROW 8
		'----------------------------------------------------
		strHTMLBody = strHTMLBody & _
		"<tr>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" & 2 * LineH & """>" & _
				"<input id='PASSWD2' name=Password2 style=""text-align: center;font-size: " & nFontSize_12 & ".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 &_
				"; border-radius: 10px " &_
				"; background-color: " & InputBGColor & "; font-weight: Normal;"" AccessKey=p size=20 maxlength=32 tabindex=3 " & _
				"type=password onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"" > " &_
			"</td>" &_
			"<td style=""border-style: none; background-color: none;"" align=""center"" class=""oa1"" height=""" &  2 * LineH & """>" & _
			"</td>" &_				
		"</tr>"
	End If
	strHTMLBody = strHTMLBody & "</tbody></table>"
    strHTMLBody = strHTMLBody &_
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = "Change Password"
	g_objIE.document.getElementById("OK").style.borderRadius = "10px"
	g_objIE.document.getElementById("EXIT").style.borderRadius = "10px"
	g_objIE.document.getElementById("OK").style.backgroundcolor = ButtonColor
	g_objIE.document.getElementById("EXIT").style.backgroundcolor = ButtonColor
	If Confirm Then
	    g_objIE.Document.getElementById("OK").innerHTML = "OK"
	Else 
	   	g_objIE.Document.getElementById("OK").innerHTML = "SIGN IN"
	End If
	
	g_objIE.Visible = False
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy	
	Set g_objShell = WScript.CreateObject("WScript.Shell")
	Call IE_Unhide(g_objIE)
	Call IE_GetPID(strPID, g_objIE.document.Title & " - " & IE_Window_Title, nDebug)
	g_objShell.AppActivate strPID
	g_objIE.Document.All("OldPass").Focus
	g_objIE.Document.All("OldPass").Value = ""
	Do
		On Error Resume Next
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case strNothing
			Case "Cancel"
				' The user clicked Cancel. Exit the loop
				IE_ChangePassword = False				
				Exit Do
			Case "OK"
				Select Case Confirm
					Case True
						if g_objIE.Document.All("Password").Value = g_objIE.Document.All("Password2").Value  and _
						   InStr(g_objIE.Document.All("Password").Value," ") = 0 and g_objIE.Document.All("Password").Value <> "" Then 
							strPassword = g_objIE.Document.All("OldPass").Value
							strNewPass = g_objIE.Document.All("Password").Value
							IE_ChangePassword = True
							Exit Do
						Else
							strNewPass = "NO PASSWORD"
							strPassword = "DO NOT MATCH"
							IE_ChangePassword = True
							Exit Do
						End If 
					Case Else
							strPassword = g_objIE.Document.All("OldPass").Value
							strNewPass = g_objIE.Document.All("Password").Value
							IE_ChangePassword = True
							Exit Do
				End Select
		End Select
	    Wscript.Sleep 200
    Loop
	g_objIE.quit
	Wscript.Sleep 200
	Set g_objIE = Nothing
	Set g_objShell = Nothing
End Function
'##############################################################################
'      Function RETURN SUCCESS AND DISPLAYS LEFT TIME
'##############################################################################
 Function IE_MSG (objParentWin, vIE_Scale, strTitle, vLine, ByVal nLine)
    Dim intX
    Dim intY
	Dim WindowH, WindowW, CellH, CellW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim nDebug
	Dim g_objIE, g_objShell
	Dim strPID
	nDebug = 0
    Set g_objShell = WScript.CreateObject("WScript.Shell")	
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	Call IE_Hide(objParentWin)
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  g_objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	IE_MSG = True
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	nTab = Round(20 * nRatioX,0)
	BottomH = Round(10 * nRatioY,0)
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
    '----------------------------------------
	'   CREATE IEXPLORER OBJECT FOR THE FORM
	'----------------------------------------
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	MainTextColor = HttpTextColor1
	Call Set_IE_obj (g_objIE)
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " 
	g_objIE.Document.body.Style.background = BackGroundColor 'HttpBgColor1
	g_objIE.Document.body.Style.color = BackGroundColor 'HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	strHTMLBody = "<br>"
	
	For nInd = 0 to nLine - 1
	    If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
		strHTMLBody = strHTMLBody &_
						"<b><p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p></b>" 
			
	Next		
	
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: None; background-color: " & ButtonColor & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & Int((CellW - nButtonX)/2) & "px; bottom: 4px' id='OK' name='OK' AccessKey='O' onclick=document.all('ButtonHandler').value='OK';><u>O</u>K</button>" & _
				"<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
'				"<input style=""position: absolute; bottom: -50px; background-color: " & BackGroundColor & """ id='KeyHandler' name='KeyHandler' value='Nothing' onkeydown=""if (event.keyCode == 13) document.all('ButtonHandler').value='OK'"">" &_				
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.document.getElementById("OK").style.borderRadius = "25px"
	g_objIE.Visible = False
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	Call IE_Unhide(g_objIE)
	WScript.Sleep 100	
	Call IE_GetPID(strPID, g_objIE.document.Title & " - " & IE_Window_Title, nDebug)
	g_objShell.AppActivate strPID
'	g_objIE.Document.All("KeyHandler").Focus
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "OK"
				IE_MSG = True
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
		g_objIE.quit
		Wscript.Sleep 200
		Set g_objIE = Nothing
		Set g_objShell = Nothing
		Call IE_Unhide(objParentWin)
End Function
'-----------------------------------------------------------------------------------------
'      Function Displays a Message with Continue and No Button. Returns True if Continue
'-----------------------------------------------------------------------------------------
 Function IE_CONT (objParentWin,vIE_Scale, strTitle, vLine, ByVal nLine, nDebug)
    Dim intX
    Dim intY
	Dim WindowH, WindowW, IE_Window_Title
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim g_objIE, g_objShell
	Set g_objShell = WScript.CreateObject("WScript.Shell")	
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
	IE_CONT = False
	Call IE_Hide(objParentWin)
	'-----------------------------------------------------------------
	'  GET THE TITLE NAME USED BY IE EXPLORER WINDOW
	'-----------------------------------------------------------------
	On Error Resume Next
		Err.Clear
		IE_Window_Title =  g_objShell.RegRead(IE_REG_KEY)
		if Err.Number <> 0 Then 
			IE_Window_Title = "Internet Explorer"
		End If
	On Error Goto 0
	'----------------------------------------
	' SCREEN RESOLUTION
	'----------------------------------------
	intX = 1920
	intY = 1080
	intX = vIE_Scale(0,2) : IE_Border = vIE_Scale(0,1) : intY = vIE_Scale(1,2) : IE_Menu_Bar = vIE_Scale(1,1)
	nRatioX = vIE_Scale(0,0)/1920
    nRatioY = vIE_Scale(1,0)/1080
	CellW = Round(350 * nRatioX,0)
	CellH = Round((150 + nLine * 30) * nRatioY,0)
	WindowW = CellW + IE_Border
	WindowH = CellH + IE_Menu_Bar
	nTab = Round(20 * nRatioX,0)
	BottomH = Round(10 * nRatioY,0)
	nFontSize_10 = Round(10 * nRatioY,0)
	nFontSize_12 = Round(12 * nRatioY,0)
	nFontSize_Def = Round(16 * nRatioY,0)
	nButtonX = Round(80 * nRatioX,0)
	nButtonY = Round(40 * nRatioY,0)
	If nButtonX < 50 then nButtonX = 50 End If
	If nButtonY < 30 then nButtonY = 30 End If
	Call Set_IE_obj (g_objIE)
    '----------------------------------------------
    '   MAIN COLORS OF THE FORM
    '----------------------------------------------	
	BackGroundColor = "grey"
	ButtonColor = HttpBgColor2
	MainTextColor = HttpTextColor1
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & BackGroundColor
	g_objIE.Document.body.Style.background = BackGroundColor
	g_objIE.Document.body.Style.color = BackGroundColor
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	strHTMLBody = "<br>"
	For nInd = 0 to nLine - 1
    If vLine(nInd,2) = HttpTextColor1 Then vLine(nInd,2) = MainTextColor
		strHTMLBody = strHTMLBody &_
						"<p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p>" 
			
	Next		
	
    strHTMLBody = strHTMLBody &_
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
				"; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; left: " & nTab & "px; bottom: " & BottomH &_
				"px' id='Button1' name='Continue' AccessKey='Y' onclick=document.all('ButtonHandler').value='YES';><u>Y</u>ES</button>" & _
				"<button style='font-weight: bold; border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 &_
				"; width:" & nButtonX & ";height:" & nButtonY & ";position: absolute; right: " & nTab & "px; bottom: " & BottomH &_
				"px' id='Button2' name='NO' AccessKey='N' onclick=document.all('ButtonHandler').value='NO';><u>N</u>O</button>" & _
                "<input name='ButtonHandler' type='hidden' value='Nothing Clicked Yet'>"
			
	g_objIE.Document.Body.innerHTML = strHTMLBody
	g_objIE.MenuBar = False
	g_objIE.StatusBar = False
	g_objIE.AddressBar = False
	g_objIE.Toolbar = False
	g_objIE.height = WindowH
	g_objIE.width = WindowW
	g_objIE.document.Title = strTitle
	g_objIE.Visible = False
	g_objIE.document.getElementById("Button1").style.borderRadius = "25px"
	g_objIE.document.getElementById("Button2").style.borderRadius = "25px"
	g_objIE.document.getElementById("Button1").style.backgroundcolor = ButtonColor
	g_objIE.document.getElementById("Button2").style.backgroundcolor = ButtonColor	
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	Call IE_Unhide(g_objIE)
	WScript.Sleep 100	
	Call IE_GetPID(strPID, g_objIE.document.Title & " - " & IE_Window_Title, nDebug)
	g_objShell.AppActivate strPID
	Do
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "NO"
				IE_CONT = False
				Exit Do
			Case "YES"
				IE_CONT = True
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
		Wscript.Sleep 200
		g_objIE.quit
		Set g_objIE = Nothing
		Set g_objShell = Nothing
        Call IE_Unhide(objParentWin)		
End Function
'----------------------------------------------------------
'   Function Set_IE_obj (byRef objIE)
'----------------------------------------------------------
Function Set_IE_obj (byRef objIE)
	Dim nCount
	Set_IE_obj = False
	nCount = 0
	Do 
		On Error Resume Next
		Err.Clear
		Set objIE = CreateObject("InternetExplorer.Application")
		Select Case Err.Number
			Case &H800704A6 
				wscript.sleep 1000
				nCount = nCount + 1
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				If nCount > 4 Then
					On Error goto 0
					Exit Function
				End If
			Case 0 
				Set_IE_obj = True
				On Error goto 0
				Exit Function
			Case Else 
				Call  TrDebug ("Set_IE_obj ERROR:" & Err.Number & " " & Err.Description, "", objDebug, MAX_LEN, 1, 1)
				On Error goto 0
				Exit Function
		End Select
	On Error goto 0
	Loop
End Function 
'-------------------------------------------------------------
'    Function WriteScreenResolution(vIE_Scale, intX,intY)
'-------------------------------------------------------------
Function WriteScreenResolution(ByRef vIE_Scale, nDebug)
Dim g_objIE, intX, intY, intXreal, intYreal, vScr
Dim f_objShell, vScreen(6), stdOutFile
	Set f_objShell = WScript.CreateObject("WScript.Shell")
	stdOutFile = "ks-screen.dat"
    strWork = f_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
    Redim vIE_Scale(2,3)
	nInd = 0
	Call Set_IE_obj(g_objIE)
	With g_objIE
		.Visible = False
		.Offline = True	
		.navigate "about:blank"
		Do
			WScript.Sleep 200
		Loop While g_objIE.Busy	
		.Document.Body.innerHTML = "<p>TEST</p>"
		.MenuBar = False
		.StatusBar = False
		.AddressBar = False
		.Toolbar = False		
		.Document.body.scroll = "no"
		.Document.body.Style.overflow = "hidden"
		.Document.body.Style.border = "None " & HttpBdColor1
		.Height = 100
		.Width = 100
    	OffsetX = .Width - .Document.body.clientWidth
		OffsetY = .Height - .Document.body.clientHeight
		If GetWin32Screen(".", vScr, nDebug) Then 
			 intXreal = vScr(0)
			 intYreal = vScr(1)
		Else 
			.FullScreen = True
			.navigate "about:blank"	
			 intXreal = .width
			 intYreal = .height
	    End If 
		.Quit
	End With
	If intXreal => 1440 Then intX = 1920 else intX = intXreal
	If intYreal => 900 Then intY = 1080  else intY = intYreal
	vIE_Scale(0,0) = intX : vIE_Scale(0,1) = OffsetX : vIE_Scale(0,2) = intXreal 
	vIE_Scale(1,0) = intY : vIE_Scale(1,1) = OffsetY : vIE_Scale(1,2) = intYreal
	Set g_objIE = Nothing
	vScreen(0) = "intX=" & intX          :	vScreen(1) = "intY=" & intY
	vScreen(2) = "OffsetX=" & OffsetX    :	vScreen(3) = "OffsetY=" & OffsetY
	vScreen(4) = "intXreal=" & intXreal  :	vScreen(5) = "intYreal=" & intYreal
	Call WriteArrayToFile(strWork & "\" & stdOutFile,vScreen, UBound(vScreen),1,0)
	set f_objShell = Nothing
End Function
'----------------------------------------------------------------
' Function WriteProgressToFile 
'----------------------------------------------------------------
 Function WriteProgressToFile(strProcessName, pProgress)
	Dim objProgressFile, f_objFSO, g_objShell
	Set g_objShell = WScript.CreateObject("WScript.Shell")	
	strWork = g_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	Set objProgressFile = f_objFSO.OpenTextFile(strWork & "\" & strProcessName & ".dat",ForAppending,True)
		objProgressFile.WriteLine pProgress
    objProgressFile.close
	Set f_objFSO = Nothing
	Set g_objShell = Nothing
End Function
'----------------------------------------------------------------
'   Function FocusToParentWindow(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function FocusToParentWindow(strPID)
Dim objShell
Call TrDebug ("FocusToParentWindow: RESTORE IE WINDOW:", "PID: " & strPID, objDebug, MAX_LEN, 1, 1) 
Const IE_PAUSE = 70
	Set objShell = WScript.CreateObject("WScript.Shell")
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "%"
	wscript.sleep IE_PAUSE
	objShell.AppActivate strPID			
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "% "
	wscript.sleep IE_PAUSE  
	objShell.SendKeys "r"
	Set objShell = Nothing
End Function
'----------------------------------------------------------------------------------------
'   Function IE_UnHide(objIE) changes the visibility of the Window referenced by the objIE
'----------------------------------------------------------------------------------------
Function IE_Unhide(byRef objIE)
    If Not IsObject(objIE) Then Exit Function
    If objIE = Null then exit function
    If Not objIE.Visible then 
        objIE.Visible = True
    End If 
End Function
'----------------------------------------------------------------------------------------
'   Function IE_Hide(objIE) changes the visibility of the Window referenced by the objIE
'----------------------------------------------------------------------------------------
Function IE_Hide(byRef objIE)
    If Not IsObject(objIE) Then Exit Function
    If objIE = Null then exit function
    If objIE.Visible then 
        objIE.Visible = False
    End If 
End Function
'----------------------------------------------------------------
'   Function IE_GetPID(strPID) Returns focus to the parent Window/Form
'----------------------------------------------------------------
Function IE_GetPID(ByRef strPID, strWinTitle, nDebug)
Dim strLine, strCmd, vCmdOut
    IE_GetPID = False
	strCmd = "tasklist /fo csv /fi ""Windowtitle eq " & strWinTitle & """"
	Call RunCmd("127.0.0.1", "", vCmdOut, strCMD, Null,nDebug)
    strPID = ""
	For Each strLine in vCmdOut
	   If InStr(strLine,"iexplore.exe") then 
	        strPID = Split(strLine,""",""")(1)
			IE_GetPID = True
			Exit For
	        Call TrDebug("IE_GetPID IE Window (" & strWinTitle & ") PID: " & strPID  , "", objDebug, MAX_LEN, 1, nDebug)
		End If
    Next
End Function
'-------------------------------------------------------
'  Function GetWin32Screen(strComputer, vScreen, nDebug) GetScreen Resolution from WMI 
'-------------------------------------------------------
Function GetWin32Screen(strComputer, ByRef vScreen, nDebug)
Dim objWMIService, colItems, objItem
GetWin32Screen = False
Redim vScreen(2)
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
	Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_VideoController") 
	On Error Resume Next
	For Each objItem in colItems
		If Err.Number <> 0 Then Call TrDebug( "Screen Resolution Error ", "", objDebug, MAX_LEN , 1, 1) : Exit For : End If	
		Call TrDebug( "Screen Resolution: " & objItem.DeviceID & " Hight : " & objItem.CurrentVerticalResolution & "  Width: " & objItem.CurrentHorizontalResolution, "", objDebug, MAX_LEN , 1, nDebug)
		If objItem.CurrentHorizontalResolution = "" or objItem.CurrentVerticalResolution = "" Then Exit For : End If
		vScreen(0) = objItem.CurrentHorizontalResolution
		vScreen(1) = objItem.CurrentVerticalResolution
		GetWin32Screen = True
		Exit For
	Next
	On Error Goto 0
End Function
'------------------------------------------------------
' Function LaunchScript
'------------------------------------------------------
Function LaunchScript(strScript,ByRef vInventory, nInventory, strARG, bWaitOnReturn)
Dim nIndex, strLaunch
Dim objShell
LaunchScript = False
	set objShell = WScript.CreateObject("WScript.Shell")
	nIndex = GetObjectLineNumber( vInventory, nInventory, strScript, True)
	If nIndex > 0 Then
		strLaunch = "wscript " & strDirectoryWork & "\" & Split(vInventory(nIndex - 1),",")(1) & "\" & Split(vInventory(nIndex - 1),",")(0)
		if strARG <> "" Then strLaunch = strLaunch & " " & strARG
		objShell.run strLaunch, 1, bWaitOnReturn
		LaunchScript = True
	End If
	set ObjShell = Nothing
End Function
'------------------------------------------------------
' Function AdminLaunchScript Launch script as administrator
'------------------------------------------------------
Function AdminLaunchScript(strScript,ByRef vInventory, nInventory, strARG)
Dim nIndex, strLaunch
Dim objShellApp
	Set objShellApp = CreateObject("Shell.Application")
	AdminLaunchScript = False
	nIndex = GetObjectLineNumber( vInventory, nInventory, strScript, True)
	If nIndex > 0 Then
		strLaunch = strDirectoryWork & "\" & Split(vInventory(nIndex - 1),",")(1) & "\" & Split(vInventory(nIndex - 1),",")(0)
		if strARG <> "" Then strLaunch = strLaunch & " " & strARG
		objShellApp.ShellExecute "wscript", strLaunch, "", "runas", 1
		AdminLaunchScript = True
	End If
	set objShellApp = Nothing
End Function
'----------------------------------------------------------------------------------
'    Function GetScreenUserSYS
'----------------------------------------------------------------------------------
Function GetScreenUserSYS()
Dim vLine
Dim strScreenUser, strUserProfile
Dim nCount
Dim objEnvar
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	strUserProfile = objEnvar.ExpandEnvironmentStrings("%USERPROFILE%")
	vLine = Split(strUserProfile,"\")
	nCount = Ubound(vLine)
	strScreenUser = vLine(nCount)
	If InStr(strScreenUser,".") <> 0 then strScreenUser = Split(strScreenUser,".")(0) End If
	set objEnvar = Nothing
	GetScreenUserSYS = strScreenUser
End Function
'----------------------------------------------------------------------------------
'    Function GetHostNameSYS
'----------------------------------------------------------------------------------
Function GetHostNameSYS()
	Dim objEnvar
	Set objEnvar = WScript.CreateObject("WScript.Shell")	
	GetHostNameSYS = objEnvar.ExpandEnvironmentStrings("%COMPUTERNAME%")
	set objEnvar = Nothing
End Function
'-----------------------------------------------------------------------
'   Function CleanFolder(strFolder, strFilePattern, strFileExtention)
'-----------------------------------------------------------------------
Function CleanFolder(strFolder, strFilePattern, strFileExtention, TimeAgeMunutes, nDebug)
    Dim objFSO,	colFiles, strFile, objFolder
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFSO.GetFolder(strFolder)
    Set colFiles = objFolder.Files
	On Error Resume Next
		For Each objFile in colFiles
			strFile = objFile.Name
			Call TrDebug( "Found File: " & strFile & "  Last Modified Date: " & DateDiff("h", objFile.DateLastModified, Date() & " " & Time()), "", objDebug, MAX_LEN , 1, nDebug)
			If (InStr(strFile,strFilePattern) <> 0 or strFilePattern = "*") and (InStr(strFile,"." & strFileExtention) <> 0 or strFileExtention = "*" ) and DateDiff("n", objFile.DateLastModified, Date() & " " & Time()) => TimeAgeMunutes Then 
			   objFSO.DeleteFile strFolder & "\" & strFile, False
			   Call TrDebug( "File: " & strFile, "Deleted", objDebug, MAX_LEN , 1, 1)
			End If
		Next
	On Error Goto 0 
	Set objFSO = Nothing
	Set objFolder = Nothing
    Set colFiles = Nothing
End Function
'--------------------------------------------------------------
' Function GetDefaultGwWMI(strComputer,nDebug)
'--------------------------------------------------------------
Function GetDefaultGwWMI(strComputer, nDebug)
Dim objWMIService, IPConfigSet, IPConfig, nGw, strGw , nMetric
    nMetric = 10000
	GetDefaultGwWMI = "0"
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Set IPConfigSet = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
	On Error resume next
	For Each IPConfig in IPConfigSet
		' strLine = strLine &	"<option value=" & IPConfig.IPAddress(0) & ">" & IPConfig.Description & "</option>"
		nGw = 0
        For each strGw in IPConfig.DefaultIPGateway
			If Err.Number > 0 Then 
				Call TrDebug ("GetDefaultGwWMI: CAN'T GET Default Gateway. Error: " & Err.Description, "ERROR" , objDebug, MAX_LEN, 1, 1)
				Err.Clear
				Exit For 
			End If
    		Call TrDebug ("Default Gateway:" & strGw & " Metrc: " & IPConfig.GatewayCostMetric(nGw), "" , objDebug, MAX_LEN, 1, nDebug)
			If IPConfig.GatewayCostMetric(nGw) < nMetric Then 
			   GetDefaultGwWMI = strGw
			   nMetric = IPConfig.GatewayCostMetric(nGw)
			End If 
			nGw = nGw + 1
		Next
	    nAdapter = nAdapter + 1	
	Next
    On Error goto 0
    Call TrDebug ("BEST GATEWAY:" & GetDefaultGwWMI & " Best Metric: " & nMetric, "" , objDebug, MAX_LEN, 1, nDebug)
End Function 
'-----------------------------------------------------------------------
' Function Function GetCmdPing(ByRef strIP, nDebug)		
'-----------------------------------------------------------------------
Function GetCmdPing(ByRef strIP, nDebug)		
    Dim  nResult, vStdOut
	Dim f_objShell
	Const SEND_PING = "ping -w 100 -n 2 "
	'--------------------------------------------------------------------------------
    '  Start Telnet session to Ping Device
    '--------------------------------------------------------------------------------
    Set f_objShell = CreateObject("WScript.Shell")
	nResult = "....."
	If RunCmd("127.0.0.1", "", vStdOut, SEND_PING & strIP, f_objShell, nDebug) = 0 Then 
	    Exit Function
	End If
	For i = 0 to UBound(vStdOut) - 1
	   If Len(vStdOut(i)) > 0 Then 
	        If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": CMD PING: " & vStdOut(i)  End If
	        If Split(vStdOut(i)," ")(0) = "Reply" Then 
			    If Split(vStdOut(i)," ")(3) = "bytes=32" Then 
			        nResult = "!!!!!"
				    Exit For
				End If
		    End If
		End If
	Next
    objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": CMD PING: " & strIP & ": Return: " & nResult  
	GetCmdPing = nResult
	Set f_objShell = Nothing
	nResult = 0
End Function
'--------------------------------------------------------------------
' Function Runs MS CMD Command on local or remote PC
'--------------------------------------------------------------------
Function RunCmd(strHost, strPsExeFolder, ByRef vCmdOut, strCMD, strAny, nDebug)	
	Dim nResult
	Dim nCmd, stdOutFile, objCmdFile, cmdFile, f_objShell
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")
	Set f_objShell = WScript.CreateObject("WScript.Shell")
	strRnd = My_Random(1,999999)
	stdOutFile = "svc-" & strRnd & ".dat"
	cmdFile = "run-" & strRnd & ".bat"
    strWork = f_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	If strHost = f_objShell.ExpandEnvironmentStrings("%COMPUTERNAME%") or strHost = "127.0.0.1" Then 
		strPsExec = ""
	Else 
		strPsExec = strPsExeFolder & "\psexec \\" & strHost & " -s "
	End If
	'-------------------------------------------------------------------
	'       CREATE A NEW TERMINAL SESSION IF REQUIRED
	'-------------------------------------------------------------------
	Set objCmdFile = objFSO.OpenTextFile(strWork & "\" & cmdFile,ForWriting,True)
	Call TrDebug ("COMMAND: ", strPsExec & strCMD & " >" & strWork & "\" & stdOutFile, objDebug, MAX_WIDTH, 1, nDebug)
	objCmdFile.WriteLine strPsExec & strCMD & " >" & strWork & "\" & stdOutFile
	objCmdFile.WriteLine "Exit"
	objCmdFile.close
	f_objShell.run strWork & "\" & cmdFile,0,True
	Call TrDebug ("BATCH FILE EXECUTED: ", strWork & "\" & cmdFile, objDebug, MAX_WIDTH, 1, nDebug)
	wscript.sleep 100
	'-----------------------------------------
	' READ OUTPUT FILE AND DELETE WHEN DONE
	'-----------------------------------------
	RunCmd = GetFileLineCountSelect(strWork & "\" & stdOutFile, vCmdOut,"NULL","NULL","NULL",0)
	If f_objFSO.FileExists(strWork & "\" & stdOutFile) Then
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & stdOutFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",stdOutFile, objDebug, MAX_WIDTH, 1, 1)
			On Error goto 0
		End If	
	End If
	If f_objFSO.FileExists(strWork & "\" & cmdFile) Then 
		On Error Resume Next
		Err.Clear
		f_objFSO.DeleteFile strWork & "\" & cmdFile, True
 		If Err.Number <> 0 Then 
			Call TrDebug ("RunCmd: ERROR CAN'T DELET FILE:",cmdFile, objDebug, MAX_WIDTH, 1, 1)
			On Error goto 0
		End If		
	End If
	Set f_objFSO = Nothing
	Set f_objShell = Nothing
	If RunCmd = 0 Then 
		Call TrDebug ("RunCmd: " & strCMD & " ERROR: ", "CAN'T WRITE TO OUTPUT FILE OR EMPTY FILE" , objDebug, MAX_WIDTH, 1, 1)
		Exit Function 
	End If
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
'---------------------------------------------------------------------------------------
'   Search for the First Line which contains "strFind" and Replaces whole Line with "strNewLine"
'   Function FindAndReplaceStrInFile(strFile, strFind, strNewLine, nDebug)
'---------------------------------------------------------------------------------------
Function FindAndReplaceStrInFile(strFile, strFind, strNewLine, nDebug)
	Dim strFolderTmp, nFileLine
	Dim vFileLine, vvFileLine
	Const FOR_WRITING = 1
	FindAndReplaceStrInFile = False
	nFileLine = GetFileLineCountSelect(strFile,vFileLine,"NULL","NULL","NULL",0)                  ' - ATTANTION nFileLine is number of lines counted like 1,2,...,n
	LineNumber = GetObjectLineNumber( vFileLine, nFileLine, strFind, True)
	If LineNumber = 0 Then 
		objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": FindAndReplaceStrInFile: ERROR CAN'T FIND """ & strFind & """ in file: " & strFile
		FindAndReplaceStrInFile = False
		Exit Function
	End If
	If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) &  ": FindAndReplaceStrInFile: LineNumber=" & LineNumber & " nFileLine=" & nFileLine  End If  
	vFileLine(LineNumber - 1) = strNewLine
	If WriteArrayToFile(strFile,vFileLine,nFileLine,FOR_WRITING,nDebug) Then FindAndReplaceStrInFile = True End If
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
'-----------------------------------------------------	
' Set nHidden to False: Make file UnHidden
' Set nHidden to True: Make file Hidden
' Function FileHidden(strFile, nHidden)
'-----------------------------------------------------
Function FileHidden(strFile, nHidden)
	Dim objFSO, objFile
	Const Hidden = 2
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strFile) Then 
		Set objFile = objFSO.GetFile(strFile)
		Select Case nHidden
			Case True
				If objFile.Attributes AND Hidden Then 
					FileHidden = True
				Else 
					objFile.Attributes = objFile.Attributes XOR Hidden
					FileHidden = True
				End If 
			Case False
				If objFile.Attributes AND Hidden Then 
					objFile.Attributes = objFile.Attributes XOR Hidden
					FileHidden = True
				Else 
					FileHidden = True
				End If 
		End Select		
	End If
	Set objFile = Nothing
	Set objFSO = Nothing
End Function
'-----------------------------------------------------------------
'     Function GetMyDate()
'-----------------------------------------------------------------
Function GetMyDate()
	GetMyDate = Month(Date()) & "/" & Day(Date()) & "/" & Year(Date()) 
End Function
'--------------------------------------------------------------
' Function returns a random intiger between min and max
'--------------------------------------------------------------
Function My_Random(min, max)
	Randomize
	My_Random = (Int((max-min+1)*Rnd+min))
End Function
'###################################################################################
' Displays a Message Box with Cancel / Continue buttons                 
'###################################################################################
Function Continue(strMsg, strTitle)
    ' Set the buttons as Yes and No, with the default button
    ' to the second button ("No", in this example)
    nButtons = vbYesNo + vbDefaultButton2
    
    ' Set the icon of the dialog to be a question mark
    nIcon = vbQuestion
    
    ' Display the dialog and set the return value of our
    ' function accordingly
    If MsgBox(strMsg, nButtons + nIcon, strTitle) <> vbYes Then
        Continue = False
    Else
        Continue = True
    End If
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
'-------------------------------------------------------
'  Function SetFileSignature(strFile)
'-------------------------------------------------------
Function SetFileSignature(strSourceFile)
	Dim objFSO,strFileString,objDestinationFile,objSourceFile,MD5cs
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = CreateObject("WScript.Shell")
	Const ForAppending = 8
	Const ForWriting = 2
	Const ForReading = 1
	Set objSourceFile = objFSO.OpenTextFile(strSourceFile,ForReading,True)
	strFileString = objSourceFile.ReadAll
'	strFileString = Split(strFileString,chr(13) & "#Signature:")(0)
    If InStr(strFileString, "#Signature:") > 0 Then
        strFileString = Split(strFileString,"#Signature:")(0)	
	    strFileString = Left(strFileString,Len(strFileString) - 2)
		'MsgBox "Set Len = " & Len(strFileString)
	End If 
	objSourceFile.close
	MD5cs = MD5Hash(strFileString)
	On Error Resume Next
	    Err.Clear
	    Set objDestinationFile = objFSO.OpenTextFile(strSourceFile,ForWriting,True)
		Select Case Err.Number
			Case 0
				objDestinationFile.WriteLine strFileString
				objDestinationFile.WriteLine "#Signature:" & MD5cs
             	objDestinationFile.close
				Set objDestinationFile = Nothing
				SetFileSignature = True
			Case Else 
			    Err.Clear
			    SetFileSignature = False
		End Select
	On Error Goto 0
	Set objSourceFile = Nothing
End Function
'--------------------------------------------------------------------------------
Function MD5Hash(sMessage)
    Dim x
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim d
    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21
    x = ConvertToWordArray(sMessage)
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476
    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d
        FFHash a, b, c, d, x(k + 0), S11, &HD76AA478
        FFHash d, a, b, c, x(k + 1), S12, &HE8C7B756
        FFHash c, d, a, b, x(k + 2), S13, &H242070DB
        FFHash b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        FFHash a, b, c, d, x(k + 4), S11, &HF57C0FAF
        FFHash d, a, b, c, x(k + 5), S12, &H4787C62A
        FFHash c, d, a, b, x(k + 6), S13, &HA8304613
        FFHash b, c, d, a, x(k + 7), S14, &HFD469501
        FFHash a, b, c, d, x(k + 8), S11, &H698098D8
        FFHash d, a, b, c, x(k + 9), S12, &H8B44F7AF
        FFHash c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        FFHash b, c, d, a, x(k + 11), S14, &H895CD7BE
        FFHash a, b, c, d, x(k + 12), S11, &H6B901122
        FFHash d, a, b, c, x(k + 13), S12, &HFD987193
        FFHash c, d, a, b, x(k + 14), S13, &HA679438E
        FFHash b, c, d, a, x(k + 15), S14, &H49B40821
        GGHash a, b, c, d, x(k + 1), S21, &HF61E2562
        GGHash d, a, b, c, x(k + 6), S22, &HC040B340
        GGHash c, d, a, b, x(k + 11), S23, &H265E5A51
        GGHash b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        GGHash a, b, c, d, x(k + 5), S21, &HD62F105D
        GGHash d, a, b, c, x(k + 10), S22, &H2441453
        GGHash c, d, a, b, x(k + 15), S23, &HD8A1E681
        GGHash b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        GGHash a, b, c, d, x(k + 9), S21, &H21E1CDE6
        GGHash d, a, b, c, x(k + 14), S22, &HC33707D6
        GGHash c, d, a, b, x(k + 3), S23, &HF4D50D87
        GGHash b, c, d, a, x(k + 8), S24, &H455A14ED
        GGHash a, b, c, d, x(k + 13), S21, &HA9E3E905
        GGHash d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        GGHash c, d, a, b, x(k + 7), S23, &H676F02D9
        GGHash b, c, d, a, x(k + 12), S24, &H8D2A4C8A
        HHHash a, b, c, d, x(k + 5), S31, &HFFFA3942
        HHHash d, a, b, c, x(k + 8), S32, &H8771F681
        HHHash c, d, a, b, x(k + 11), S33, &H6D9D6122
        HHHash b, c, d, a, x(k + 14), S34, &HFDE5380C
        HHHash a, b, c, d, x(k + 1), S31, &HA4BEEA44
        HHHash d, a, b, c, x(k + 4), S32, &H4BDECFA9
        HHHash c, d, a, b, x(k + 7), S33, &HF6BB4B60
        HHHash b, c, d, a, x(k + 10), S34, &HBEBFBC70
        HHHash a, b, c, d, x(k + 13), S31, &H289B7EC6
        HHHash d, a, b, c, x(k + 0), S32, &HEAA127FA
        HHHash c, d, a, b, x(k + 3), S33, &HD4EF3085
        HHHash b, c, d, a, x(k + 6), S34, &H4881D05
        HHHash a, b, c, d, x(k + 9), S31, &HD9D4D039
        HHHash d, a, b, c, x(k + 12), S32, &HE6DB99E5
        HHHash c, d, a, b, x(k + 15), S33, &H1FA27CF8
        HHHash b, c, d, a, x(k + 2), S34, &HC4AC5665
        IIHash a, b, c, d, x(k + 0), S41, &HF4292244
        IIHash d, a, b, c, x(k + 7), S42, &H432AFF97
        IIHash c, d, a, b, x(k + 14), S43, &HAB9423A7
        IIHash b, c, d, a, x(k + 5), S44, &HFC93A039
        IIHash a, b, c, d, x(k + 12), S41, &H655B59C3
        IIHash d, a, b, c, x(k + 3), S42, &H8F0CCC92
        IIHash c, d, a, b, x(k + 10), S43, &HFFEFF47D
        IIHash b, c, d, a, x(k + 1), S44, &H85845DD1
        IIHash a, b, c, d, x(k + 8), S41, &H6FA87E4F
        IIHash d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        IIHash c, d, a, b, x(k + 6), S43, &HA3014314
        IIHash b, c, d, a, x(k + 13), S44, &H4E0811A1
        IIHash a, b, c, d, x(k + 4), S41, &HF7537E82
        IIHash d, a, b, c, x(k + 11), S42, &HBD3AF235
        IIHash c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        IIHash b, c, d, a, x(k + 9), S44, &HEB86D391
        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next
    MD5Hash = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
'--------------------------------------------------------------------------------
Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
    AddUnsigned = lResult
End Function
'--------------------------------------------------------------------------------
Function FFHash(Byref a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(FHash(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Function
'--------------------------------------------------------------------------------
Function GGHash(ByRef a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(GHash(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Function
'--------------------------------------------------------------------------------
Function HHHash(ByRef a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(HHash(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Function
'--------------------------------------------------------------------------------
Function IIHash(ByRef a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(IHash(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Function
'--------------------------------------------------------------------------------
Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    Const BITS_TO_A_BYTE = 8
    Const BYTES_TO_A_WORD = 4
    Const BITS_TO_A_WORD = 32
    lMessageLength = Len(sMessage)
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop
    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    ConvertToWordArray = lWordArray
End Function
'--------------------------------------------------------------------------------
Function WordToHex(lValue)
    Dim lByte
    Dim lCount
    Const BITS_TO_A_BYTE = 8
    Const BYTES_TO_A_WORD = 4
    Const BITS_TO_A_WORD = 32
    Dim m_lOnBits(30)
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)	
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function
'--------------------------------------------------------------------------------
Function LShift(lValue, iShiftBits)
    Dim m_l2Power(30)
	Dim m_lOnBits(30)
	m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)	
	
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function
'--------------------------------------------------------------------------------
Function RShift(lValue, iShiftBits)
    Dim m_l2Power(30)
	m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function
'--------------------------------------------------------------------------------
Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function
'--------------------------------------------------------------------------------
Function FHash(x, y, z)
    FHash = (x And y) Or ((Not x) And z)
End Function
'--------------------------------------------------------------------------------
Function GHash(x, y, z)
    GHash = (x And z) Or (y And (Not z))
End Function
'--------------------------------------------------------------------------------
Function HHash(x, y, z)
    HHash = (x Xor y Xor z)
End Function
'--------------------------------------------------------------------------------
Function IHash(x, y, z)
    IHash = (y Xor (x Or (Not z)))
End Function
