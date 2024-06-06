'Option Explicit
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , chr(34)&WScript.ScriptFullName&chr(34) & " /elevate", "", "runas", 1
  WScript.Quit
End If

Dim objShell: Set objShell = CreateObject("WScript.Shell")

Select Case objShell.Popup("This script will attempt to install excelUNFORS, and download and Autoit"&vbCrLf&"Please ensure all files are extracted to the folder you wish to install to."&vbCrLf&vbCrLf&"Do you wish to continue?", 10, "Install excelUNFORS", 1)
Case -1 
    'Timed Out
    WScript.Quit
Case 1
    Install("")
Case 2
    WScript.Quit
End Select


Function GetPath(ByVal argumentName)
    GetPath= CreateObject("WScript.Shell").ExpandEnvironmentStrings(argumentName)
End Function



Function Install(ByVal argumentName)

Dim scriptDir, filepath
scriptDir = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\") - 1)

Dim objShell, objShortcut, strFolder

Set objShell = WScript.CreateObject("WScript.Shell")
strFolder = objShell.SpecialFolders("AllUsersStartMenu")
Set objShortcut = objShell.CreateShortcut(strFolder & "\ExcelUNFORS.lnk")
objShortcut.TargetPath = ScriptDir&"\excelunfors.au3"
objShortcut.WorkingDirectory = ScriptDir
objShortcut.WindowStyle = 1
objShortcut.Hotkey = "CTRL+SHIFT+F12"
objShortcut.Description = "Dosimeter Excel Entry Software that searches for all cells containing ***, and higlights them"
objShortcut.Save
Set objShortcut = Nothing
Set objShell = Nothing

filepath=GetPath("%USERPROFILE%") & "\Downloads\"&"autoit-v3-setup.zip"

DIM fso    
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(filepath)=False Then
 Set fso = CreateObject("Scripting.FileSystemObject")
 Set Outp = Wscript.Stdout
 On Error Resume Next
 Set File = WScript.CreateObject("Microsoft.XMLHTTP")
 File.Open "GET", "https://www.autoitscript.com/cgi-bin/getfile.pl?autoit3/autoit-v3-setup.zip", False
 File.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 1.1.4322; .NET CLR 3.5.30729; .NET CLR 3.0.30618; .NET4.0C; .NET4.0E; BCD2000; BCD2000)"
 File.Send
 If err.number <> 0 then 
  Outp.writeline "" 
  Outp.writeline "Error getting file" 
  Outp.writeline "==================" 
  Outp.writeline "" 
  Outp.writeline "Error " & err.number & "(0x" & hex(err.number) & ") " & err.description 
  Outp.writeline "Source " & err.source 
  Outp.writeline "" 
  Outp.writeline "HTTP Error " & File.Status & " " & File.StatusText
  Outp.writeline  File.getAllResponseHeaders
  Outp.writeline Arg(1)
 End If

On Error Goto 0

 Set BS = CreateObject("ADODB.Stream")
 BS.type = 1
 BS.open
 BS.Write File.ResponseBody

 BS.SaveToFile filepath, 2

End If

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/C" &" "&chr(34)&filepath&chr(34)

If (fso.FileExists(filepath)) Then
  WScript.Echo("Download successful, please install Autoit"&vbCrLf&vbCrLf&"excelUNFORS.au3 icon added to start menu"&vbCrLf)
Else
  WScript.Echo("Some part of the install failed, maybe all files where not extracted to the same folder, or no internet?"&vbCrLf)
End If

End Function