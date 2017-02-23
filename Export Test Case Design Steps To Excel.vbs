<html>
<HEAD>
<TITLE>Test Case Design Steps Export Utility</TITLE>
<style type="text/css">
	form, fieldset, h2, table{
		margin: 0px;
	}
	body {
		background-color: #BDB76B;
		text-align: center;
	}
	#DomainList, #ProjectList, #UserNameText, #PasswordText {
		width: 160px;
	}
	#ServerNameText, #ExportFrom, #ExportTo {
		width: 220px;
	}

</style>
<script type="text/vbscript">
Const ForReading = 1
Dim FsoIn
Dim TDConnection
Dim cmd
Dim rec
Dim GUIState
Dim LastDBName
Dim LastDomain
Dim TestCaseCount, AttachmentCount
Dim objFSO, objProgressMsg, oWsh, WScript
Dim logStream
Dim ExportNonDesignStepFlag, ShowProgressFlag, ExportLogFlag, DebugFlag, ExportAttachmentsFlag, ExportFolderAttachmentsFlag
Dim strWindowTitle

strWindowTitle = "Test Case Design Steps Export Utility"
DebugFlag = False
ExportLogFlag = False
ExportNonDesignStepFlag = False
ShowProgressFlag = False

Set oWsh = CreateObject("WScript.Shell")
Set TDConnection = CreateObject ("TDApiOle80.TDConnection.1")
Set WScript = CreateObject("WScript.Shell")

Sub PopulateDomainList
	On Error Resume Next
	DomainList.onchange= ""
	DomainList.innerHTML = ""
	Err.Clear

	For i = 1 to TDConnection.VisibleDomains.Count
		Domain = TDConnection.VisibleDomains.Item(i)
		If Err <> 0 Then Exit For
		Set thisOpt = document.createElement( "option" )
		thisOpt.innerHTML = Domain
		thisOpt.value = Domain
		DomainList.appendChild thisOpt
	Next
	If Err <> 0 Then
		ProjectList.innerHTML = ""
		window.alert "Cannot list domains. Check User credentials." & vbCrLf & Err.Description
	Else
		Call PopulateProjectList()
	End if
End Sub

Sub PopulateProjectList

	On Error Resume Next
	ProjectList.innerHTML = ""
	Err.Clear

	For i=1 to TDConnection.VisibleProjects(DomainList.value).Count
		If Err <> 0 Then Exit For
		Set thisOpt = document.createElement( "option" )
		Project = TDConnection.VisibleProjects(DomainList.value).Item(i) 'IG
		thisOpt.innerHTML = Project
		thisOpt.value = Project
		If Project = LastProject Then thisOpt.selected = True
		ProjectList.appendChild thisOpt
	Next
	If Err <> 0 Then window.alert "Cannot list projects" & vbCrLf & Err.Description
End Sub

Sub setGUIState(newState)
	On Error Resume Next
	GUIState = Ucase(newState)
	Select Case GUIState
		Case "SERVER_DISCONNECTED"
			ServerNameText.disabled = false
			UserNameText.disabled = false
			PasswordText.disabled = false
			ConnectButton.disabled = false
			ConnectButton.innerText = "Connect"
			DomainProject.style.display = "none"
			FolderSelect.style.display = "none"
			Config.style.display = "none"
			Results.style.display = "none"
		Case "SERVER_CONNECTED"
			ConnectButton.disabled = false
			ConnectButton.innerText = "Disconnect"
			ServerNameText.disabled = true
			UserNameText.disabled = true
			PasswordText.disabled = true
			DomainProject.style.display = "block"
			FolderSelect.style.display = "block"
			Config.style.display = "block"
			Results.style.display = "block"
			Call PopulateDomainList
		Case "PROJECT_CONNECTED"
			DomainProject.style.display = "none"
			FolderSelect.style.display = "none"
			Config.style.display = "none"
			Results.style.display = "block"
		Case "PROJECT_DISCONNECTED"
			DomainProject.style.display = "block"
			FolderSelect.style.display = "block"
			Config.style.display = "block"
			Results.style.display = "block"
		Case Else
			' unknown type, ignoring
	End Select
	window.status = "Ready"

	On Error GoTo 0

End Sub

Sub Report (Msg)
	Dim oSpan
	Set oSpan = document.createElement ("SPAN")
	oSpan.innerHTML = "<BR>" & CStr (Now()) & " : " & Msg
	ExportLog Msg
	Results.appendChild(oSpan)'.scrollIntoView
	Set oSpan = Nothing
	If ShowProgressFlag Then ProgressMsg Msg, strWindowTitle
End Sub

Sub ExportLog(Msg)
	If ExportLogFlag Then
		logStream.writeline CStr(Now()) & " : " & Msg
	End If
End Sub

Sub ReportError (sFunction, sMsg)
	window.alert "Error in function " & sFunction & vbCrLf & sMsg
End Sub

Sub Disconnect
	On Error Resume Next
	Err.Clear
	If TDConnection.Connected Then
		If TDConnection.ProjectConnected Then TDConnection.DisconnectProject
		TDConnection.ReleaseConnection
	End If
	If Err <> 0 Then ReportError "Disconnect", Err.Description
	setGUIState "SERVER_DISCONNECTED"
End Sub

Function trimDLLName(ServerName)
	Dim newName
	newName = ServerName
	If InStr (LCase (newName), "/wdomsrv.dll") <> 0 Then
		newName = Mid(newName, 1, Len(newName) - Len("/wdomsrv.dll"))
	End If
	If InStr (LCase (newName), "/wcomsrv.dll") <> 0 Then
		newName = Mid(newName, 1, Len(newName) - Len("/wcomsrv.dll"))
	End If
	If InStr (LCase (newName), "/tdservlet") <> 0 Then
		newName = Mid(newName, 1, Len(newName) - Len("/tdservlet"))
	End If
	If InStr (LCase (newName), "/servlet") <> 0 Then
		newName = Mid(newName, 1, Len(newName) - Len("/servlet"))
	End If

	trimDLLName = newName
End Function


Sub ExportFolder(TestCasePath,ExportPath)
	If DebugFlag Then msgbox TestCasePath & " " & ExportPath
	TCFolderPath = TestCasePath
	If Right(TCFolderPath,1) = "\" Then
		TCFolderPath = Left(TCFolderPath, len(TCFolderPath)-1)
	End If
	If Right(ExportPath,1) = "\" Then
		ExportPath = Left(ExportPath, len(ExportPath)-1)
	End If
	Report "Exporting test cases from: " & TestCasePath
	Report "Exporting test cases to: " & ExportPath
	Set TestF = TDConnection.TestFactory

	Set TreeMgr = TDConnection.TreeManager
	Set CurrentTreeNode = TreeMgr.NodeByPath("Subject\" & TCFolderPath)

	Set TestFilter = TestF.Filter
	TestFilter.Filter("TS_SUBJECT") = Chr(34) & "Subject\" & TCFolderPath & chr(34)

	Set TestList = TestF.NewList(TestFilter.Text)

	If TestList.Count >= 1 Then
		ExportLog "There are " & TestList.Count & " possible tests to export. " & "Subject\" & TCFolderPath
		For i = 1 to TestList.Count
			If TestList.Item(i).DesStepsNum > 0 or ExportNonDesignStepFlag = True Then
				Call ExportDStoExcel(TestList.Item(i),ExportPath)
				TestCaseCount = TestCaseCount + 1
			Else
				ExportLog "This test case has no design steps. Skipping " & TestList.Item(i).Name
			End If
		Next
	End If

	'If current folder has subfolders
	'for each subfolder call Exportfolder function recursively
	If CurrentTreeNode.Count >= 1 Then
		ExportLog "There are " & CurrentTreeNode.Count & " subfolders. " & "Subject\" & TCFolderPath
		For m = 1 to CurrentTreeNode.Count
			ExportFolder TCFolderPath & "\" & CurrentTreeNode.Child(m).Name,ExportPath& "\" & CurrentTreeNode.Child(m).Name
		Next
	End If
End Sub


Sub ExportDStoExcel(objTest,ExportPath)

	If Right(ExportPath,1) = "\" Then
		ExportPath = Left(ExportPath, len(ExportPath)-1)
	End If

	Set DesStepsF = objTest.DesignStepFactory
	Set DesStepList = DesStepsF.NewList("")
	If DesStepList.Count = 0 and ExportNonDesignStepFlag = False Then
		'No design steps present
		ExportLog "This test case has no design steps. Skipping " & objTest.Name
		Exit Sub
	End If

	Set objExcelApp = CreateObject("Excel.Application")
	objExcelApp.DisplayAlerts = False
	If chkShowExcel.Checked Then
		objExcelApp.Visible = True
	Else
		objExcelApp.Visible = False
	End If
	Set ExcelWorkbook = objExcelApp.Workbooks.Add
	Set ExcelSheet = ExcelWorkbook.Worksheets(1)
	ExcelWorkbook.Worksheets("Sheet2").Delete
	ExcelWorkbook.Worksheets("Sheet3").Delete

	ExcelSheet.Name = "Test Case"
	ExcelSheet.Cells(1, 1).Value = "Test Case:"
	ExcelSheet.Cells(2, 1).Value = "Test Description:"

	ExcelSheet.Columns("A:A").Font.Bold = True
	ExcelSheet.Columns("A:A").Interior.ColorIndex = 15
	ExcelSheet.Columns("A:A").Interior.Pattern = 1

	DesStepsRow = 4
	If DesStepList.Count <> 0 Then ExcelSheet.Cells(DesStepsRow, 1).Value = "Step Name"
	If DesStepList.Count <> 0 Then ExcelSheet.Cells(DesStepsRow, 2).Value = "Step Description"
	If DesStepList.Count <> 0 Then ExcelSheet.Cells(DesStepsRow, 3).Value = "Expected Result"
	If DesStepList.Count <> 0 Then ExcelSheet.Cells(DesStepsRow, 4).Value = "Actual Result"

	If DesStepList.Count <> 0 Then ExcelSheet.Rows(DesStepsRow & ":" & DesStepsRow).Font.Bold = True
	If DesStepList.Count <> 0 Then ExcelSheet.Range("B" & DesStepsRow & ":D" & DesStepsRow).Interior.ColorIndex = 15
	If DesStepList.Count <> 0 Then ExcelSheet.Range("B" & DesStepsRow & ":D" & DesStepsRow).Interior.Pattern = 1

	ExcelSheet.Columns.AutoFit
	ExcelSheet.Columns("B:B").ColumnWidth = 50

	ExcelSheet.Cells(1, 2).Value = objTest.Name
	ExcelSheet.Cells(2, 2).Value = ConvertHTMLDisplayTags(ReplaceText("<[^<>]*>", "",objTest.Field("TS_Description")))

	ExcelSheet.Cells.WrapText = True

	For j = 1 to DesStepList.Count
		ExcelSheet.Cells(j+DesStepsRow, 1).Value = ConvertHTMLDisplayTags(DesStepList.Item(j).StepName)
		ExcelSheet.Cells(j+DesStepsRow, 2).Value = ConvertHTMLDisplayTags(ReplaceText("<[^<>]*>", "",DesStepList.Item(j).StepDescription))
		ExcelSheet.Cells(j+DesStepsRow, 3).Value = ConvertHTMLDisplayTags(ReplaceText("<[^<>]*>", "",DesStepList.Item(j).StepExpectedResult))
	Next

	Set fso = CreateObject("Scripting.FileSystemObject")
	If Not fso.FolderExists(ExportPath) Then
		EPathArray = Split(ExportPath,"\",-1,1)
		EPathString = EPathArray(0)
		For k = 1 to UBound(ePathArray)
			EPathString = EPathString & "\" & EPathArray(k)
			If Not fso.FolderExists(EPathString) Then
				Set f = fso.CreateFolder(EPathString)
				Report "Creating folder: " & f.Path
			End If
		Next
	End If

	ExcelWorkbook.SaveAs ExportPath & "\" & objTest.Name & ".xlsx"
	ExportLog "Spreadsheet created and saved. " & objTest.Name & ".xlsx"
	ExcelWorkbook.Close
	objExcelApp.Quit
	Set objExcelApp = Nothing
	Set ExcelWorkbook = Nothing
	Set ExcelSheet = Nothing
End Sub

Function ReplaceText(byVal patrn, byVal replStr, byVal textString)
	Dim regEx, Matches, Match           ' Create variables.
	ReplaceText = textString
	Set regEx = New RegExp            ' Create regular expression.
	regEx.Pattern = patrn            ' Set pattern.
	'regEx.IgnoreCase = True            ' Make case insensitive.
	regEx.Global = True         ' Set global applicability.

	Set Matches = regEx.Execute(textString)
	For Each Match in Matches
		ReplaceText = regEx.Replace(textString, replStr)   ' Make replacement.
	Next
End Function

Function ConvertHTMLDisplayTags(htmlText)
	htmlText = ReplaceText("&lt;","<",htmlText)
	htmlText = ReplaceText("&gt;",">",htmlText)
    htmlText = ReplaceText("&quot;",chr(34),htmlText)
	ConvertHTMLDisplayTags = htmlText
End Function


Function ProgressMsg( strMessage, strWindowTitle )
' Obtained from Internet Forums
' Written by Denis St-Pierre
' Displays a progress message box that the originating script can kill in both 2k and XP
' If StrMessage is blank, take down previous progress message box
' Using 4096 in Msgbox below makes the progress message float on top of things
' CAVEAT: You must have   Dim ObjProgressMsg   at the top of your script for this to work as described
    'Set wshShell = WScript.CreateObject( "WScript.Shell" )
    Set wshShell = CreateObject( "WScript.Shell" )
    strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
	On Error Resume Next
	' Kill ProgressMsg
	objProgressMsg.Terminate( )
	' Re-enable Error Checking
	On Error Goto 0

    If strMessage = "" Then
		Exit Function
    End If
    Set objFSOMSG = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs"     'Control File for reboot

    ' Create Message.vbs, True=overwrite
    'msgbox "prior to createtext file"
    On error resume next
    Set objTempMessage = objFSOMSG.CreateTextFile( strTempVBS, True )
'    If err then
'    	WScript.Sleep 180
'    	Set objTempMessage = objFSOMSG.CreateTextFile( strTempVBS, True )
'	End If
	on error goto 0
    objTempMessage.WriteLine( "MsgBox""" & strMessage & """, 4096, """ & strWindowTitle & """" )
    objTempMessage.Close

    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kills the Previous ProgressMsg
    objProgressMsg.Terminate( )
    ' Re-enable Error Checking
    On Error Goto 0

    ' Trigger objProgressMsg and keep an object on it
    Set objProgressMsg = wshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )

    Set wshShell = Nothing
    Set objFSOMSG = Nothing
End Function
</script>


<script FOR="ConnectButton" EVENT="onclick()" type="text/vbscript">
	On Error Resume Next
	If GUIState = "SERVER_DISCONNECTED"  Then
		window.status = "Connecting to server..."
		Err.Clear
		TDConnection.InitConnectionEx ServerNameText.value
		If Err <> 0 Then
			window.alert "Cannot connect to  '" & ServerNameText.value & "'" & vbCrLf & Err.Description
			setGUIState "Server_Disconnected"
			Beep
		Else
			TDConnection.Login UserNameText.value,PasswordText.value
			If err = 0 and TDConnection.Connected Then
				setGUIState "Server_Connected"
			Else
				setGUIState "Server_Disconnected"
				Beep
				window.alert "Cannot login to  '" & ServerNameText.value & "'. Check Login credentials." & vbCrLf & Err.Description
			End If
		End If
	Else
		Disconnect
		setGUIState "Server_Disconnected"
	End If
	On Error GoTo 0
</script>
<script FOR="ExportButton" EVENT="onclick()" type="text/vbscript">
	If ExportFrom.value = "" or ExportTo.value = "" Then
		Report "The Export To field and the Export From fields are required."
	Else
		window.status = "Exporting Design Steps..."
		TestCaseCount = 0
		AttachmentCount = 0
		TDConnection.ConnectProjectEx DomainList.value, ProjectList.value , UserNameText.value , PasswordText.value
		If Err = 0 Then
			Results.innerHTML = ""
			If chkExportLog.Checked = True Then
				ExportLogFlag = True
				Set objFSO = CreateObject("scripting.filesystemobject")
				Set logStream = objFSO.createtextfile(ExportTo.value & "\Export.log" , True)
			Else
				ExportLogFlag = False
			End If
			If chkIncludeNonDesignStepTests.Checked = True Then
				ExportNonDesignStepFlag = True
			Else
				ExportNonDesignStepFlag = False
			End If

			If chkShowProgress.Checked = True Then
				ShowProgressFlag = True
			Else
				ShowProgressFlag = False
			End If	

			Report "Connected to project " & DomainList.value & "." & ProjectList.value
			setGUIState "Project_Connected"

			ExportFolder ExportFrom.value, ExportTo.value

			TDConnection.DisconnectProject
			setGUIState "Project_Disconnected"
			If AttachmentCount > 0 then
				Report "Test Cases Exported: " & TestCaseCount & "; Attachments Exported: " & AttachmentCount
			else
				Report "Test Cases Exported: " & TestCaseCount
			End if
			If ExportLogFlag Then logStream.Close
		Else
			window.status = Err.Description
			Report "Error in Exporting while connecting to QC Project: " & Err.Description
			Beep
		End If
	End If
	On Error Goto 0
</script>
<script FOR="RefreshProjectButton" EVENT="onclick()" type="text/vbscript">
	On Error Resume Next
	PopulateProjectList
	On Error GoTo 0
</script>
<script FOR="DomainList" EVENT="onclick()" type="text/vbscript">
	On Error Resume Next
	PopulateProjectList
	On Error GoTo 0
</script>
<script FOR="DomainList" EVENT="onchange()" type="text/vbscript">
	PopulateProjectList
</script>

<script FOR="BrowseToFolderButton" EVENT="onclick()" type="text/vbscript">
	On Error Resume Next
        dim objShell
        dim ssfWINDOWS
        dim objFolder
		'ssfWINDOWS = 36
        set objShell = CreateObject("Shell.Application")
            set objFolder = objShell.BrowseForFolder(0, "Select Folder", 0)
			if (not objFolder is nothing) then
				ExportTo.value = objFolder.Items.Item.Path
				'ExportTo.value = objFolderItem.Path
			end if
            set objFolder = nothing
        set objShell = nothing
	On Error GoTo 0
</script>
<script type="text/vbscript">
Sub OnLoadBody
	Dim IniFile
	Dim oWshEnv
	Dim LastServer
	Dim LastUser
	Dim thisOpt
	Dim i
	Dim objNet
	On Error Resume Next
	Set oWsh = CreateObject("WScript.Shell")
	LastServer = oWsh.RegRead ("HKCU\SOFTWARE\Mercury Interactive\TestDirector\WEB\LastConnection")
	ServerNameText.value = trimDLLName (LastServer)

	Set objNet = CreateObject("WScript.NetWork")
    UserNameText.value = objNet.UserName
	Set objNet = Nothing

	Set oWshEnv = oWsh.Environment ("Process")
	IniFile = CStr (oWshEnv("SystemRoot")) & "\mercury.ini"
	Set oWshEnv = Nothing
	Set FsoIn = CreateObject ("Scripting.FileSystemObject")
	Set FsoIn = Nothing
	setGUIState "SERVER_DISCONNECTED"
	On Error GoTo 0
End Sub

Sub OnUnloadBody
	Call Disconnect
	Set oWsh = Nothing
	Set TDConnection = Nothing
End Sub
</script>
</HEAD>
<body onload="OnLoadBody" onunload="OnUnloadBody">
	<table>
	<caption><h3>Designed By : Thirupathi Gangula
		<tr valign="middle">
			<td align="left" valign="middle">
			<H2>Test Case Design Steps Export Utility
			</td>
			
		</tr>
	</table>
	<hr>
	<table>
		<caption><h3>Enter Quality Center URL, UserID and Password
		<tr>
			<td align=left>
				<label for="ServerNameText"><B>Quality Center URL:</B></label>
			</td>
			<td align=left>
				<input id="ServerNameText">
			</td>
		</tr>
		<tr>
			<td align=left>
				<label for="UserNameText"><B>Username:</B></label>
			</td>
			<td align=left>
				<input type="username" id="UserNameText">
			</td>
		</tr>
		<tr>
			<td align=left>
				<label for="PasswordText"><B>Password:</B></label>
			</td>
			<td align=left>
				<input type="password" id="PasswordText" >
			</td>
		</tr>
	</table>
	<button id="ConnectButton"><B>Connect</B></button>
	<div id="DomainProject" align=center>
		<hr>
		<table>
			<caption><h3>Select QC Domain and Project
			<tr>
				<td align=left>
					<label for="DomainList"><B>Domain:</B></label>
				</td>
				<td align=left>
					<select id="DomainList">
						<option selected>[None]</option>
					</select>
				</td>
			</tr>
			<tr>
				<td align=left>
					<label for="ProjectList"><B>Project:</B></label>
				</td>
				<td align=left>
					<select id="ProjectList">
						<option selected>[None]</option>
					</select>
					<button id="RefreshProjectButton">Refresh</button>
				</td>
			</tr>
		</table>
	</div>
	<div id="Config" align=center>
		<hr>
		<table>
			<caption><h3>Select Configuration Options<BR>
			<tr>
				<td align=left>
					<input type=checkbox name="chkShowProgress" checked value="ON">Show Progress Messages
				<td align=left>
					<input type=checkbox name="chkShowExcel" value="OFF">Show Excel During Creation
			</tr>
			<tr>
				<td align=left>
					<input type=checkbox name="chkIncludeNonDesignStepTests" checked value="ON">Include Tests With No Design Steps
				<td align=left>
					<input type=checkbox name="chkExportLog" checked value="ON">Create Log File
			</tr>
		</table>
	</div>
	<div id="FolderSelect" align=center>
		<hr>
		<table>
			<caption><h3>Choose QC Folder to Export From and Location to Export To
			<tr>
				<td align=left>
					<label for="ExportFrom"><B>Export From:</B></label>
				</td>
				<td align=left>
					Subject\<input id="ExportFrom">
				</td>
			</tr>
			<tr>
				<td align=left>
					<label for="ExportTo"><B>Export To:</B></label>
				</td>
				<td align=left>
					<input id="ExportTo" disabled><button id="BrowseToFolderButton">...</button>
				</td>
			</tr>
			<tr>
				<td align=center colspan=2>
					<button id="ExportButton">Export Design Steps</button>
				</td>
			</tr>
			</table>
	</div>
	<hr>
	<div id="Results" align=left>
	</div>
</body>
</html>
