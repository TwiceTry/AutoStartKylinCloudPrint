' 获取当前脚本所在目录
Set objFSO = CreateObject("Scripting.FileSystemObject")
currentDirectory = objFSO.GetAbsolutePathName(".")
ps1FilePath = ""

' 查找第一个.ps1脚本
Set folder = objFSO.GetFolder(currentDirectory)
Set files = folder.Files
For Each file In files
    If LCase(objFSO.GetExtensionName(file.Name)) = "ps1" Then
        ps1FilePath = file.Path
        Exit For
    End If
Next

If ps1FilePath = "" Then
    WScript.Echo "未找到.ps1脚本"
    WScript.Quit
End If

Set wshShell = CreateObject("WScript.Shell")

' 提示用户输入密码
password = InputBox("请输入服务端密码:", "设置密码")

' 定义生成的VBS脚本名称（统一管理）
scriptName = "RunKylinPrint"

' 定义日志文件路径
logFile = currentDirectory & "\" & scriptName & ".log"

' 定义生成的VBS脚本内容
generatedVBSContent = _
"Set ws = WScript.CreateObject(""WScript.Shell"")" & vbCrLf & _
"ws.Run ""powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & Chr(34) & ps1FilePath & Chr(34) & Chr(34) & " -password " & Chr(34) & Chr(34) & password & Chr(34) & Chr(34) & " -WorkingDirectory " & Chr(34) &  Chr(34) &  Replace(currentDirectory, "\", "\\") & Chr(34) & Chr(34) &  " > " & Chr(34) & Chr(34) & logFile & Chr(34) & Chr(34) & " 2>&1"", 0, False"

' 生成的VBS脚本路径
generatedVBSScriptPath = currentDirectory & "\" & scriptName & ".vbs"

' 创建或覆盖已有的VBS脚本
Set generatedFile = objFSO.CreateTextFile(generatedVBSScriptPath, True)
generatedFile.WriteLine generatedVBSContent
generatedFile.Close

' 添加注册表项设置开机启动
Const REGISTRY_PATH = "Software\Microsoft\Windows\CurrentVersion\Run"
On Error Resume Next
wshShell.RegWrite "HKCU\" & REGISTRY_PATH & "\" & scriptName, generatedVBSScriptPath, "REG_SZ"
If Err.Number = 0 Then
    WScript.Echo "启动项已成功添加！"
Else
    WScript.Echo "添加启动项时出错: " & Err.Description
End If
On Error GoTo 0

' 创建桌面快捷方式到生成的VBS脚本
desktopPath = wshShell.SpecialFolders("Desktop")
shortcutPath = desktopPath & "\" & scriptName & ".lnk"
' 定义图标来源的.exe文件路径
iconSourceExe = currentDirectory & "\kylin-cloud-printer-server.exe"
' 创建快捷方式
Set shortcut = wshShell.CreateShortcut(shortcutPath)
shortcut.TargetPath = Chr(34) & generatedVBSScriptPath & Chr(34)
shortcut.Description = "运行自动麒麟云打印"
shortcut.WorkingDirectory = currentDirectory
shortcut.IconLocation = iconSourceExe & ",0" ' 第二个参数是图标的索引，默认为0
shortcut.Save

WScript.Echo "已生成桌面快捷方式"