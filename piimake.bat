<comment rem>
for %%x in (system32 syswow64) do if exist "%SystemRoot%\%%x" set SystemLeaf=%%x
start "%~n0" "%SystemRoot%\%SystemLeaf%\mshta.exe" "%~f0"
goto :eof
</comment>
<head>
<title>HTML Application host for Windows CE PII Creation Assistant</title>
<meta http-equiv="MSThemeCompatible" content="yes">
<style>
body
{
	margin: 5px 5px 60px 5px;
	font: 14px sans-serif;
	background-color: silver;
	overflow: hidden;
	border: none;
}
fieldset
{
	height: 100%;
}
fieldset iframe
{
	width: 24.9%;
	height: 50%;
	zoom: 75%;
}
div
{
	white-space: nowrap;
}
div label
{
	padding-right: 10px;
}
a
{
	color: blue;
	background-color: silver;
	position: absolute;
	right: 15px;
}
span
{
	position: absolute;
	right: 10px;
	bottom: 5px;
}
</style>
<script type="text/vbs">
Option Explicit

Const AddOnName = "HTML Application host for Windows CE"

SetLocale 1033

Dim fso, wsh
Set fso = CreateObject("Scripting.FileSystemObject")
Set wsh = CreateObject("WScript.Shell")

Dim home, inst
home = fso.GetParentFolderName(location.pathname)
inst = wsh.RegRead("HKCR\CLSID\{A31E2E44-714B-11D6-8A19-000102228262}\LocalServer32\")
inst = fso.GetParentFolderName(Replace(inst, """", ""))

Function IsAdmin
	On Error Resume Next
	wsh.RegRead "HKEY_USERS\S-1-5-19\Environment\TEMP"
	IsAdmin = Err.number = 0
End Function

Function AddOnFolder
	AddOnFolder = Replace(home, home, inst, 1, Intrusive.checked) & "\AddOn"
End Function

Function CreateFolder(path)
	On Error Resume Next
	fso.CreateFolder path
	CreateFolder = Err.Number = 0
End Function

Function DeleteFolder(path)
	On Error Resume Next
	fso.DeleteFolder path
	DeleteFolder = Err.Number = 0
End Function

Sub CreateAddon_OnClick
	Dim i, frame, line, path, file
	CreateFolder(AddOnFolder)
	If CreateFolder(AddOnFolder & "\" & AddOnName) Then DeleteAddon.disabled = False
	path = AddOnFolder & "\" & AddOnName & "\Common"
	If CreateFolder(path) Then
		fso.CopyFile home & "\show-vbs-info.bat", path & "\"
	End If
	For i = 0 To document.frames.length - 1
		Set frame = document.frames(i)
		path = AddOnFolder & "\" & AddOnName & "\" & frame.frameElement.name
		If CreateFolder(path) Then
			fso.CopyFile home & "\" & frame.frameElement.title & "\Release\cehta.exe", path & "\"
		End If
		path = AddOnFolder & "\" & AddOnName & "\" & fso.GetFileName(frame.frameElement.src)
		Set file = fso.CreateTextFile(path, True)
		For Each line In Split(frame.document.body.innerText, vbCrLf)
			line = Trim(line)
			if Len(line) > 4 And InStr(line, "#name") = Len(line) - 4 Then
				file.WriteLine AddOnName & "#name"
			ElseIf Len(line) > 21 And InStr(line, "#TARGET_os_version_") = Len(line) - 21 Then
				file.WriteLine FormatNumber(Right(frame.frameElement.name, 3) / 100, 2) & " " & Right(line, 22)
			ElseIf InStr(1, line, "; file ", vbTextCompare) = 1 Then
				file.WriteLine "\Common\show-vbs-info.bat > \flash\AddOn\ #NO"
				file.WriteLine "\" & frame.frameElement.name & "\cehta.exe > \flash\AddOn\ #NO"
			' ElseIf InStr(1, line, "; registry ", vbTextCompare) = 1 Then
			' ElseIf InStr(1, line, "; uninstall ", vbTextCompare) = 1 Then
			ElseIf Len(line) <> 0 And InStr(line, "\") = 0 And InStr(line, ";") = 0 Then
				file.WriteLine line
			End If
		Next
	Next
End Sub

Sub DeleteAddon_OnClick
	If DeleteFolder(AddOnFolder & "\" & AddOnName) Then DeleteAddon.disabled = True
End Sub

Sub Intrusive_OnClick
	DeleteAddon.disabled = Not fso.FolderExists(AddOnFolder & "\" & AddOnName)
End Sub

Sub ShowLicense_OnClick
	showModalDialog "LICENSE", Nothing, "dialogWidth=41em"
End Sub

Sub Window_OnLoad
	Dim i, frame
	For i = 0 To document.frames.length - 1
		Set frame = document.frames(i)
		frame.frameElement.src = Replace(frame.frameElement.src, "about:", inst & "\AddOn\HTML_AddOn\")
	Next
	Intrusive.disabled = Not IsAdmin
	Intrusive.checked = Not Intrusive.disabled
	DeleteAddon.disabled = Not fso.FolderExists(AddOnFolder & "\" & AddOnName)
	Version.innerText = fso.GetFileVersion(home & "\Compact2013_SDK_86Duino_80B\Release\cehta.exe")
End Sub
</script>
</head>
<body>
<a href="#" unselectable="on" onclick="vbs:wsh.Run(Me.innerText)">https://github.com/datadiode/cehta</a>
<span disabled id="Version"></span>
<fieldset>
<legend>Templates</legend>
<iframe name="arm_800" title="WEC2013 Beaglebone SDK" src="about:KTP_Mob_4.pii"></iframe>
<iframe name="arm_800" title="WEC2013 Beaglebone SDK" src="about:KTP_Mobile_7_9.pii"></iframe>
<iframe name="arm_800" title="WEC2013 Beaglebone SDK" src="about:TP_10F_Mobile.pii"></iframe>
<iframe name="arm_600" title="Beckhoff_HMI_600 (ARMV4I)" src="about:CP_4.pii"></iframe>
<iframe name="x86_600" title="Beckhoff_HMI_600 (x86)" src="about:CP_7_9.pii"></iframe>
<iframe name="x86_600" title="Beckhoff_HMI_600 (x86)" src="about:CP_7_15_Out.pii"></iframe>
<iframe name="x86_600" title="Beckhoff_HMI_600 (x86)" src="about:CP_15.pii"></iframe>
<iframe name="x86_800" title="Compact2013_SDK_86Duino_80B" src="about:CP_GX_800.pii"></iframe>
</fieldset>
<div>
<button id="CreateAddon">Create ProSave Addon</button>
<button id="DeleteAddon">Delete ProSave Addon</button>
<label for="Intrusive" title="This option allows an install right beside ProSave's stock addons (requires admin rights)">
<input id="Intrusive" type="checkbox">Intrusive</label>
<button id="ShowLicense">&#9878; Show License</button>
</div>
</body>
