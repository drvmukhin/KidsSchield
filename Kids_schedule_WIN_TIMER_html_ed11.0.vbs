'----------------------------------------------------------------------------------
'            KIDS SCHILD TIMER SCRIPT VERSION 11.0
'----------------------------------------------------------------------------------
	Dim strDirectoryLCL
strDirectoryLCL = "C:\Users\All Users\Vandyke"
' KEEP strSrvDirectory VARIABLE AT LINE 5 TO LET AUTOINSTALATION
'----------------------------------------------------------------------------------
Const ForAppending = 8
Const ForWriting = 2
Const HttpTextColor1 = "#292626"
Const HttpTextColor2 = "#F0BC1F"
Const HttpTextColor3 = "#EBEAF7"
Const HttpTextColor4 = "#A4A4A4"
Const HttpBgColor1 = "Grey"
Const HttpBgColor2 = "#292626"
Const HttpBgColor3 = "#0404B4"
Const HttpBgColor4 = "#504E4E"
Const HttpBgColor5 = "#0D057F"
Const HttpBgColor6 = "#8B9091"
Const HttpBdColor1 = "Grey"
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
Const PAUSE_ACCOUNT = 	"$Pause"
Const BLOCK_ACCOUNT = 	"$Block"
Const BLOCK_TIMER = 	"BLOCK"
Const MAX_TIMER = 	20
Dim nResult
Dim strMonthMaxFileName, strFileAccount, strFileDevice, strFileString
Dim strFileSession
Dim strSrvDirectory, strDirectoryUpdate, strDirectoryWork, strDirectoryVandyke, strCRTexe
Dim strDeviceID
Dim strAccountID
Dim nDebug
Dim vActivity
Dim nDevice, nAccount, nActivity, nSession, nSessionTmp
Dim vDevice, vAccount, vSession, vvMsg(20,3)
Dim vLine
Dim objFSO, objDebug, objShell
Dim vConnect
Dim strNetworkStatus, strServerStatus, vIE_Scale
strFileSession = "sessions.txt"
strFileSessionTmp = "sessions_tmp.dat"
strSrvDirectory = "\\MEDIA\_PublicFolder\KidsSchild"
strDirectoryWork = "C:\KidsSchild"
strDirectoryUpdate = "\\HIGS\Install_My\Tools_Networking\KidsSchild"
strDirectoryVandyke = "C:\Program Files"
strCRTexe = """C:\Program Files\VanDyke Software\SecureCRT\SecureCRT.exe"""
strFileWinUserTmp = "winusers_tmp.dat"
strFileConnect = "connectivity_tmp.dat"
strVersion = "None"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = WScript.CreateObject("WScript.Shell")
nDebug = 0
Main()
  
If IsObject(objDebug) Then objDebug.Close : End If
set objFSO = Nothing
Set objShell = Nothing
Sub Main()
'-----------------------------------------------------------------
'  GET SCREEN RESOLUTION
'-----------------------------------------------------------------
	Call GetScreenResolution(vIE_Scale, nDebug)
'-----------------------------------------------------------------
'  Check if TIMER script is already running
'-----------------------------------------------------------------
	On Error Resume Next
	Set objDebug = objFSO.OpenTextFile(strDirectoryLCL & "\" & "debug-timer.log",ForAppending,True)
	Select Case Err.Number
		Case 0
		Case 70
			vvMsg(0,0) = "TIMER IS ALREADY RUNNING" 			    	: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
			vvMsg(1,0) = "Exit . . ."									: vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2 
 			Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,2)
			Exit Sub
		Case Else 
			vvMsg(0,0) = "CAN'T LAUNCH TIMER"	 						: vvMsg(0,1) = "normal" : vvMsg(0,2) = "Red"
			vvMsg(1,0) = "Error# " & Err.Number & " " & Err.Source   	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1
			vvMsg(2,0) = "Error# " & Err.Description					: vvMsg(2,1) = "normal" : vvMsg(2,2) = HttpTextColor1
			vvMsg(3,0) = "Exit . . ."									: vvMsg(3,1) = "bold" 	: vvMsg(3,2) = HttpTextColor2
 			Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,4)
			Exit Sub
	End Select
	On Error goto 0
	If objFSO.FileExists(strDirectoryLCL & "\" & strFileSession) Then 
		nSession = GetFileLineCountSelect(strDirectoryLCL & "\" & strFileSession, vSession,"#","NULL","NULL",nDebug)
		If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": Folder: " & strFolder & " Session: " & strSession & " Host: " & strHost End If
        strSrvDirectory = vSession(1)
		If Right(strSrvDirectory,1) = "\" Then         '<----------------------------------------------- Normalize Working Directory Path
			strSrvDirectory = Left(strSrvDirectory,Len(strSrvDirectory) - 1)
			If nDebug = 1 Then objDebug.WriteLine GetMyDate() & " " & FormatDateTime(Time(), 3) & ": Remote Server Folder: " & strSrvDirectory End If
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
	strDeviceID = strLclDeviceID
	strWinUtilsFolder = strDirectoryWork & "\Bin"
	strCRTexe = """" & strDirectoryVandyke & "\SecureCRT.exe"""
	strFileAccount = strSrvDirectory & "\" & "accounts.dat"
	strFileDevice = strSrvDirectory & "\" & "devices.dat"
	strFileString = strSrvDirectory & "\" & "string.dat"
	strMonthMaxFileName = strSrvDirectory & "\" & "month-max.dat"
'-------------------------------------------------------------------------------------------
'           CHECK LOCAL PC CONNECTIVITY TO THE NETWORK AND SERVER AVAILABILITY
'-------------------------------------------------------------------------------------------
	strNetworkStatus = OFF_LINE
	strServerStatus = OFF_LINE
	
	If objFSO.FileExists(strDirectoryLCL & "\Temp\" & strFileConnect) Then 
		nConnect = GetFileLineCountSelect(strDirectoryLCL & "\Temp\" & strFileConnect, vConnect,"#", "", "",nDebug)
		strNetworkStatus = Split(vConnect(0),",")(1)
		strServerStatus = Split(vConnect(1),",")(1)
	End If
	If strNetworkStatus = OFF_LINE Then 
		vvMsg(0,0) = "COMPUTER IS NOT CONNECTED TO THE NETWORK!" 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
		vvMsg(1,0) = "Exit . . ."						: vvMsg(1,1) = "bold" : vvMsg(1,2) = HttpTextColor2 
 		Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,2)
		Exit Sub
	End If
	If strServerStatus = OFF_LINE Then 
		vvMsg(0,0) = "HOME MEDIA SERVER IS DOWN." 				: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
		vvMsg(1,0) = "TURN ON SERVER AND TRY AGAIN IN 5 Min" 	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor1			
		vvMsg(2,0) = "Exit . . ."								: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2 
 		Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,3)
		Exit Sub
	End If
	If strServerStatus = ON_LINE Then 
		vvMsg(0,0) = "HOME MEDIA SERVER IS LOADING NOW..." 		: vvMsg(0,1) = "normal" : vvMsg(0,2) = HttpTextColor1
		vvMsg(1,0) = "DO NOT TOUCH SERVER AND WAIT FOR 5 Min" 	: vvMsg(1,1) = "normal" : vvMsg(1,2) = HttpTextColor2			
		vvMsg(2,0) = "Exit . . ."								: vvMsg(2,1) = "bold" 	: vvMsg(2,2) = HttpTextColor2 
 		Call IE_MSG(g_objIE,vIE_Scale, "Error",vvMsg,3)
		Exit Sub
	End If
'-------------------------------------------------------------------------------------
'               CHECK IF ACCOUNT AND DEVICE FILES EXIST
'-------------------------------------------------------------------------------------
    If Not objFSO.FileExists(strFileAccount) Then
		MsgBox "CAN'T FIND LIST OF USERS! PLEASE TRY AGAIN, IF IT DOESN'T WORK CALL TO YOUR FATHER (650)-996-9085" 
		Exit Sub
	End If	
    If Not objFSO.FileExists(strFileDevice) Then
		MsgBox "CAN'T FIND LIST OF AVAILABLE DEVICES! PLEASE TRY AGAIN, IF IT DOESN'T WORK CALL TO YOUR FATHER (650)-996-9085" 
		Exit Sub
	End If	
	
'-------------------------------------------------------------------------------------
'               LOAD DATA FROM FILES
'-------------------------------------------------------------------------------------
	nDevice = GetFileLineCountSelect(strFileDevice, vDevice, "#", "", "",0)
	nAccount = GetFileLineCountSelect(strFileAccount, vAccount, "#", "", "",nDebug)
'-------------------------------------------------------------------------------------
'    GetActive Account and Qouta (if any) on the local device (strDeviceID)
' 	 strDeviceID is retrieved from session.txt file
'-------------------------------------------------------------------------------------
	nQuota 	= GetActiveAcctOnDevice( strSrvDirectory, strDeviceID, vAccount, nAccount, vActivity, 0) 
	If nQuota > 0 Then
		strAccountID = vActivity(2,0)
		Select Case strAccountID
			Case BLOCK_ACCOUNT
				vvMsg(0,0) = "THIS COMPUTER IS TEMPORARY LOCKED" 	:	vvMsg(0,1) = "normal" 	:	vvMsg(0,2) = HttpTextColor1
				vvMsg(1,0) = "LOCK EXPIRES IN:"				 		:	vvMsg(1,1) = "bold" 	:	vvMsg(1,2) = HttpTextColor2
				nLine = 2
			Case PAUSE_ACCOUNT
				vvMsg(0,0) = "THIS COMPUTER WAS PAUSED"		 			:	vvMsg(0,1) = "normal" 	:	vvMsg(0,2) = HttpTextColor1
				vvMsg(1,0) = "TRY TO LOCK / UNLOCK SCREEN OR REBOOT PC"	:	vvMsg(1,1) = "normal" 	:	vvMsg(1,2) = HttpTextColor1				
				vvMsg(2,0) = "PAUSE EXPIRES IN:"				 		:	vvMsg(2,1) = "bold" 	:	vvMsg(2,2) = HttpTextColor2
				nLine = 3
			Case Else
				vvMsg(0,0) = "THERE IS ONE ACTIVE SESSION CONNECTED AT LOCAL PC:" 	:	vvMsg(0,1) = "normal" 	:	vvMsg(0,2) = HttpTextColor1
				vvMsg(1,0) = "User: " & strAccountID					 		:	vvMsg(1,1) = "bold" 	:	vvMsg(1,2) = HttpTextColor1
				vvMsg(2,0) = "Time left:" 										:	vvMsg(2,1) = "bold" 	:	vvMsg(2,2) = HttpTextColor2
				nLine = 3
		End Select
		Do
         if Not IE_MSG_TIME (vIE_Scale, "Session Info", vvMsg, nLine, nQuota) then
            exit sub
        end if
	    exit Do
		Loop
	Else
		vvMsg(0,0) = "THERE ARE NO ACTIVE SESSIONS CONNECTED AT LOCAL PC" 	: vvMsg(0,1) = "normal" :	vvMsg(0,2) = HttpTextColor1
		vvMsg(1,0) = "Goodby . . ." 										: vvMsg(1,1) = "bold" 	:	vvMsg(1,2) = HttpTextColor2
         Call IE_MSG (g_objIE,vIE_Scale, "Session Info", vvMsg, 2)
	End If
End Sub
'##############################################################################
'      Function RETURN SUCCESS AND DISPLAYS LEFT TIME
'##############################################################################
 Function IE_MSG_TIME (vIE_Scale, strTitle, vLine, ByVal nLine, nQuota)
    Dim g_objIE
    Dim g_objShell
    Dim intX
    Dim intY
	Dim WindowH, WindowW
	Dim nFontSize_Def, nFontSize_10, nFontSize_12
	Dim nInd
	Dim intHH, intMM, intSS
	Dim strHH, strMM, strSS
	Dim EndTime
	Dim strQuota
	Dim nDebug
    Set g_objShell = WScript.CreateObject("WScript.Shell")	
    Dim IE_Window_Title
	Const IE_REG_KEY = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window Title"
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
	nDebug = 0
	EndTime = Timer() + nQuota * 60
	StartDate = Date()
	IE_MSG_TIME = True
	Call Set_IE_obj (g_objIE)
	CellW = Round(350 * nRatioX,0)
	CellH = Round((200 + nLine * 30) * nRatioY,0)
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
	
   If nDebug = 1 Then MsgBox "intX=" & intX & "   intY=" & intY & "   RatioX=" & nRatioX & "  RatioY=" & nRatioY & "   Cell Width=" & cellW & "  Cell Hight=" & cellH End If
	g_objIE.Offline = True
	g_objIE.navigate "about:blank"
	g_objIE.Document.body.Style.FontFamily = "Helvetica"
	g_objIE.Document.body.Style.FontSize = nFontSize_Def
	g_objIE.Document.body.scroll = "no"
	g_objIE.Document.body.Style.overflow = "hidden"
	g_objIE.Document.body.Style.border = "none " & HttpBdColor1
	g_objIE.Document.body.Style.background = HttpBgColor1
	g_objIE.Document.body.Style.color = HttpTextColor1
	g_objIE.Top = (intY - WindowH)/2
	g_objIE.Left = (intX - WindowW)/2
	strHTMLBody = "<br>"
	For nInd = 0 to nLine - 1
		strHTMLBody = strHTMLBody &_
			"<b><p style=""text-align: center; font-weight: " & vLine(nInd,1) & "; color: " & vLine(nInd,2) & """>" & vLine(nInd,0) & "</p></b>" 
			
	Next		
	strHTMLBody = strHTMLBody &_
			"<input name=UserInput style=""text-align: center; font-size: " & nFontSize_12 &_
			".0pt; border-style: None; font-family: 'Helvetica'; color: " & HttpTextColor2 & "; background-color: " &_
			HttpBgColor2 & "; font-weight: Normal;"" AccessKey=i size=40 maxlength=32 " & _
			"type=text > " &_
			"<br>"
		
    strHTMLBody = strHTMLBody &_
                "<button style='border-style: None; background-color: " & HttpBgColor2 & "; color: " & HttpTextColor2 & "; width:" & nButtonX & ";height:" & nButtonY &_
				";position: absolute; left: " & Round((CellW - nButtonX)/2,0) & "px; bottom: 4px' name='OK' AccessKey='O' onclick=document.all('ButtonHandler').value='OK';><u>O</u>K</button>" & _
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
	Do
		WScript.Sleep 100
	Loop While g_objIE.Busy
	Call IE_Unhide(g_objIE)
	g_objShell.AppActivate g_objIE.document.Title & " - " & IE_Window_Title
	Do
		If StartDate = Date() Then nQuota = EndTime - Timer Else nQuota = EndTime - Timer - 24 * 3600 End If
		intHH = Int(nQuota/3600)
		intMM = Int((nQuota - intHH * 3600)/60)
		intSS = Int(nQuota - intHH * 3600 - intMM * 60)
		If intMM < 10 Then strMM = ":0" & intMM Else strMM = ":" & intMM End If
		If intSS < 10 Then strSS = ":0" & intSS Else strSS = ":" & intSS End If
		strQuota = intHH & strMM & strSS
		If (intHH + intMM + intSS) <= 0 Then 
    			IE_MSG_TIME = True
				g_objIE.quit
				Exit Do
		End If
		On Error Resume Next
		g_objIE.Document.All("UserInput").Value = Left(strQuota,8)
		Err.Clear
		strNothing = g_objIE.Document.All("ButtonHandler").Value
		if Err.Number <> 0 then exit do
		On Error Goto 0
		Select Case g_objIE.Document.All("ButtonHandler").Value
			Case "OK"
				IE_MSG_TIME = True
				g_objIE.quit
				Set g_objIE = Nothing
				Set g_objShell = Nothing
				Exit Do
		End Select
		Wscript.Sleep 500
		Loop
End Function
'###################################################################################
'    Function GetActiveAcctOnDevice Returns Device STATUS and Looged in Account name
'###################################################################################
Function GetActiveAcctOnDevice( strDirectory, ByRef strDeviceID, ByRef vAccount, ByRef nAccount, ByRef vActiv, ByRef nDebug)
	Dim AcctInd
	Dim strAcct
	Dim strFileWeek, strLine
	Dim vFileLines
' Dim vActiv(4,100)
    GetActiveAcctOnDevice = 0
	AcctInd = 0
	strAcctID = "None"
		Do While AcctInd < nAccount
            strAcct = Split(vAccount(AcctInd),",")(0) 
            If nDebug = 1 Then objDebug.WriteLine "Get Active Accout: strAcctID=" & strAcct & "  DevInd=" & DevInd End If  
            '--------------------------------------------------------------------------------
            '                  GET WEEK and FILE NAME
            '--------------------------------------------------------------------------------
            strFileWeek = strDirectory & "\" & GetWeekFile(strAcct, nDebug)
            '--------------------------------------------------------------------------------
            ' Check IF Week file exists and create new one if it doesn't '
            '---------------------------------------------------------------------------------
            If Not objFSO.FileExists(strFileWeek) Then
                Set objDataFileWeek = objFSO.CreateTextFile(strFileWeek)
		        strLine = "0,0,0,Internet,Tablet"
		        objDataFileWeek.WriteLine strLine
		        objDataFileWeek.Close
	        End If
            '--------------------------------------------------------------------------------
            '  Read Last active devices from Account Week File 
            '--------------------------------------------------------------------------------
	        nCount = GetFileLineCountSelect(strFileWeek, vFileLines,"#","","NULL", 0)
	        nCount = GetLastActivity(nCount,strDeviceID, strAcct, vFileLines,vActiv,0)
			If nDebug = 1 Then objDebug.WriteLine "Get Active Accout: strAcctID=" & strAcct & "  DevInd=" & DevInd & " nActivity=" & vActiv(0,0) End If  
			If IsNumeric(vActiv(0,0)) and nCount > 0 Then 
						GetActiveAcctOnDevice = vActiv(0,0)
						Exit Do
			End If
		    AcctInd = AcctInd + 1
        Loop
End Function
'###################################################################################
'   GetWeekFile returns the name of the user Weekly data file for the current week
'###################################################################################
Function GetWeekFile(strAccountID, nDebug)
    Dim WeekDate, nDay
	WeekDate = DateAdd("d",-(Weekday(Date()) - 1),Date())
	If Len(Day(weekdate)) = 1 Then nDay = "0" & day(weekdate) Else nDay = day(weekdate) End If
	strFileName = strAccountID & "-" & Year(WeekDate) & Month(weekdate) & nDay & "-week.dat"
    Call TrDebug("Week Filename: " & strFileName, "", objDebug, MAX_LEN , 1, nDebug)			
    GetWeekFile = strFileName  
 End Function 
 '#######################################################################
 ' vActivity (3,n) - DeviceID
 ' vActivity (2,n) - AccountID
 ' vActivity (1,n) - ModeID
 ' vActivity (0,n) - Rest of Usage (if > 0) or GAME OVER or STOP
 ' Number of records for the account on the given device (GetLastActivity)
 ' Returns:
 ' Function GetLastActivity - Returns Last Active Seesions
 '#######################################################################
 Function GetLastActivity( ByRef nCount, ByRef strDeviceID, ByRef strAcct, ByRef vFileLines, ByRef vActivity, nDebug)
	Dim nQuota, nTime, nUsage
	Dim HH, MM
	Dim nInd,nReverse
	Dim strLine
	Dim strUserDeviceID, strDate
	Redim vActivity(4,nCount)
	nQuota = 0
	HH = Hour(Time())
	MM = Minute(Time())	
    nInd = nCount
	nReverse = 0
	Do While nInd > 0
	    strLine = vFileLines(nInd-1)
		If nDebug = 1 Then objDebug.WriteLine "GetLastActivity: vFileLines(" & nInd-1 & ") = " & vFileLines(nInd-1) End If
        strDate = Split(strLine,",")(0)
		nTime = Split(strLine,",")(1)
	    nUsage = Split(strLine,",")(2)
	    strUserDeviceID = Split(strLine,",")(4)
		If strDate <> GetMyDate() Then Exit Do '###### CHECK THE CURRENT DATE	
		If strUserDeviceID = strDeviceID  Then       '###### CHECK DEVICE NAME
			vActivity(1,nReverse) = Split(strLine,",")(3)
			vActivity(2,nReverse) = strAcct
			vActivity(3,nReverse) = strUserDeviceID
			If nUsage > 0 Then                           '###### CHECK CURRENT USAGE, IF IT IS < 0 THEN Print "STOP"
				nQuota = - HH * 60 - MM + nTime + nUsage
				If nQuota > 0 Then 
					vActivity(0,nReverse) = nQuota
					If nDebug = 1 Then objDebug.WriteLine "GetLastActivity: vActivity(" & nReverse & ") =" & vActivity(0,nReverse) End If
					nReverse = nReverse + 1
				Else
					vActivity(0,nReverse) = "GAME OVER"
					If nDebug = 1 Then objDebug.WriteLine "GetLastActivity: vActivity(" & nReverse & ") =" & vActivity(0,nReverse) End If
					nReverse = nReverse + 1
				End If	
			Else 
				vActivity(0,nReverse) = "STOP"
				If nDebug = 1 Then objDebug.WriteLine "GetLastActivity: vActivity(" & nReverse & ") =" & vActivity(0,nReverse) End If
				nReverse = nReverse + 1
			End If
		End If
        nInd = nInd - 1
    Loop
	If nDebug = 1 Then objDebug.WriteLine "GetLastActivity: nReverse=" & nReverse End If
	GetLastActivity = nReverse
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
'    Function GetScreenResolution(vIE_Scale, intX,intY)
'-------------------------------------------------------------
Function GetScreenResolution(ByRef vIE_Scale, nDebug)
Dim vScreen(6), nScreen
Dim f_objFSO, f_objShell,stdOutFile
    Redim vIE_Scale(2,3)
    GetScreenResolution = False
	Set f_objFSO = CreateObject("Scripting.FileSystemObject")	
	Set f_objShell = WScript.CreateObject("WScript.Shell")
	stdOutFile = "ks-screen.dat"
    strWork = f_objShell.ExpandEnvironmentStrings("%USERPROFILE%")
	If Not f_objFSO.FileExists(strWork & "\" & stdOutFile) Then Exit Function
    nScreen = GetFileLineCountSelect(strWork & "\" & stdOutFile,vScreen,"NULL", "NULL", "NULL",0)
	For n = 0 to nScreen - 1
	   If InStr(vScreen(n),"=") <> 0 Then vScreen(n) = Split(vScreen(n),"=")(1)
	Next 
	vIE_Scale(0,0) = vScreen(0) : vIE_Scale(1,0) = vScreen(1)
	vIE_Scale(0,1) = vScreen(2)  : vIE_Scale(1,1) = vScreen(3)
	vIE_Scale(0,2) = vScreen(4) : vIE_Scale(1,2) = vScreen(5)
	set f_objFSO = Nothing
	set f_objShell = Nothing
	GetScreenResolution = True
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
