<!DOCTYPE rem|
@echo off
if %os%==%os:/=/% goto desktop
start cehta.exe "%0"
goto :eof
:desktop
start "%~n0" "%SystemRoot%\system32\mshta.exe" "%~f0"
goto :eof
>
<?cehta-options dialogWidth=26; dialogHeight=12; resizable=yes?>
<html>
<head>
<hta:application id='dialogArguments'/>
<style>
table
{
	font-family: monospace;
	border: 1px solid black;
	border-collapse: collapse;
}
</style>
<script type='text/vbs'>
Option Explicit

Sub Window_OnLoad
	document.title = dialogArguments.commandLine & " (" & document.compatMode & ")"
	Dim row, text
	For Each row In report.rows
		text = row.cells(0).innerText
		If InStr(text, ".") = 0 Then
			row.cells(1).innerText = Eval(text)
		Else
			On Error Resume Next
			CreateObject text
			If Err.Number > 0 Then
				row.cells(1).innerText = "ERR#" & Err.Number
			ElseIf Err.Number < 0 Then
				' likely an HRESULT, therefore use hex notation
				row.cells(1).innerText = "ERR#0x" & Hex(Err.Number)
			End If
			On Error GoTo 0
		End If
	Next
End Sub
</script>
</head>
<table id='report' border='1' cellpadding='5'>
<tr><td>ScriptEngine</td><td id="tdScriptEngine"></td></tr>
<tr><td>ScriptEngineMajorVersion</td><td></td></tr>
<tr><td>ScriptEngineMinorVersion</td><td></td></tr>
<tr><td>ScriptEngineBuildVersion</td><td></td></tr>
<tr><td>VBScript.RegExp</td><td>OK</td></tr>
<tr><td>SRELL.RegExp</td><td>OK</td></tr>
</table>
</html>
