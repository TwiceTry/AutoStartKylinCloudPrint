' ��ȡ��ǰ�ű�����Ŀ¼
Set objFSO = CreateObject("Scripting.FileSystemObject")
currentDirectory = objFSO.GetAbsolutePathName(".")
ps1FilePath = ""

' ���ҵ�һ��.ps1�ű�
Set folder = objFSO.GetFolder(currentDirectory)
Set files = folder.Files
For Each file In files
    If LCase(objFSO.GetExtensionName(file.Name)) = "ps1" Then
        ps1FilePath = file.Path
        Exit For
    End If
Next

If ps1FilePath = "" Then
    WScript.Echo "δ�ҵ�.ps1�ű�"
    WScript.Quit
End If

Set wshShell = CreateObject("WScript.Shell")

' ��ʾ�û���������
password = InputBox("��������������:", "��������")

' �������ɵ�VBS�ű����ƣ�ͳһ����
scriptName = "RunKylinPrint"

' ������־�ļ�·��
logFile = currentDirectory & "\" & scriptName & ".log"

' �������ɵ�VBS�ű�����
generatedVBSContent = _
"Set ws = WScript.CreateObject(""WScript.Shell"")" & vbCrLf & _
"ws.Run ""powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File " & Chr(34) & Chr(34) & ps1FilePath & Chr(34) & Chr(34) & " -password " & Chr(34) & Chr(34) & password & Chr(34) & Chr(34) & " -WorkingDirectory " & Chr(34) &  Chr(34) &  Replace(currentDirectory, "\", "\\") & Chr(34) & Chr(34) &  " > " & Chr(34) & Chr(34) & logFile & Chr(34) & Chr(34) & " 2>&1"", 0, False"

' ���ɵ�VBS�ű�·��
generatedVBSScriptPath = currentDirectory & "\" & scriptName & ".vbs"

' �����򸲸����е�VBS�ű�
Set generatedFile = objFSO.CreateTextFile(generatedVBSScriptPath, True)
generatedFile.WriteLine generatedVBSContent
generatedFile.Close

' ���ע��������ÿ�������
Const REGISTRY_PATH = "Software\Microsoft\Windows\CurrentVersion\Run"
On Error Resume Next
wshShell.RegWrite "HKCU\" & REGISTRY_PATH & "\" & scriptName, generatedVBSScriptPath, "REG_SZ"
If Err.Number = 0 Then
    WScript.Echo "�������ѳɹ���ӣ�"
Else
    WScript.Echo "���������ʱ����: " & Err.Description
End If
On Error GoTo 0

' ���������ݷ�ʽ�����ɵ�VBS�ű�
desktopPath = wshShell.SpecialFolders("Desktop")
shortcutPath = desktopPath & "\" & scriptName & ".lnk"
' ����ͼ����Դ��.exe�ļ�·��
iconSourceExe = currentDirectory & "\kylin-cloud-printer-server.exe"
' ������ݷ�ʽ
Set shortcut = wshShell.CreateShortcut(shortcutPath)
shortcut.TargetPath = Chr(34) & generatedVBSScriptPath & Chr(34)
shortcut.Description = "�����Զ������ƴ�ӡ"
shortcut.WorkingDirectory = currentDirectory
shortcut.IconLocation = iconSourceExe & ",0" ' �ڶ���������ͼ���������Ĭ��Ϊ0
shortcut.Save

WScript.Echo "�����������ݷ�ʽ"