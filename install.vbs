'Option Explicit
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , chr(34)&WScript.ScriptFullName&chr(34) & " /elevate", "", "runas", 1
  WScript.Quit
End If

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

Set objShell = CreateObject("Shell.Application")
objShell.ShellExecute "cmd.exe", "/C" &" "&chr(34)&filepath&chr(34)
'objShell.ShellExecute """"&GetPath("%windir%") & "\system32\"&"cmd.exe /C start "&""""& filepath &""""&"""", "", "", "runas", 0
'CreateObject("Wscript.Shell").Run GetPath("%windir%") & "\system32\"&"cmd.exe /C start "&""""& filepath &"""", 1, True

Function GetPath(ByVal argumentName)
    GetPath= CreateObject("WScript.Shell").ExpandEnvironmentStrings(argumentName)
End Function
