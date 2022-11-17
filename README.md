# HTML Application host for Windows CE (cehta)
[![StandWithUkraine](https://raw.githubusercontent.com/vshymanskyy/StandWithUkraine/main/badges/StandWithUkraine.svg)](https://github.com/vshymanskyy/StandWithUkraine/blob/main/docs/README.md)

### Example of CEhta compatible .hta code wrapped into a .bat file 
```html
<!-- rem>
::^^ The XML comment brackets prevent the .bat stub from showing up on the page
@echo off
if %os%==%os:/=/% goto desktop
start cehta.exe %0
goto :eof
:desktop
start "%~n0" "%SystemRoot%\system32\mshta.exe" "%~f0"
goto :eof
-->
<html>
<![CDATA[
<!-- ^^^ The CDATA tag helps ensure XML conformance for CEhta's happiness -->
<head>
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

' Avoid Window_OnLoad as it fires twice when hosted in CEhta
Sub Document_OnReadyStateChange
	tdScriptEngine.innerText = tdScriptEngine.innerText & ScriptEngine
	tdScriptEngineMajorVersion.innerText = ScriptEngineMajorVersion
	tdScriptEngineMinorVersion.innerText = ScriptEngineMinorVersion
	tdScriptEngineBuildVersion.innerText = ScriptEngineBuildVersion
End Sub
</script>
</head>
<table border='1' cellpadding='5'>
<tr><td>ScriptEngine</td><td id="tdScriptEngine"></td></tr>
<tr><td>ScriptEngineMajorVersion</td><td id="tdScriptEngineMajorVersion"></td></tr>
<tr><td>ScriptEngineMinorVersion</td><td id="tdScriptEngineMinorVersion"></td></tr>
<tr><td>ScriptEngineBuildVersion</td><td id="tdScriptEngineBuildVersion"></td></tr>
</table>
<comment>]]>
<!-- ^^^ The <comment> tag prevents the ]]> from showing up on the page -->
</html>

```
### How it looks like in Virtual CEPC
![cehta](https://user-images.githubusercontent.com/10423465/202555698-85f0f9cb-13fb-40dc-ad1f-34bf70d60482.png)
