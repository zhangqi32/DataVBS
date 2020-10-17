Set unNamedArguments = WScript.Arguments.UnNamed 
set WshShell = WScript.CreateObject("WScript.Shell") 
Set objFSO = CreateObject("Scripting.FileSystemObject") 
' 修为自己的目标文件夹
strFolder = "C:\Users\hwzhao\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\"
for count = 0 to wscript.arguments.count-1 Step 1 
    filename = unNamedArguments.Item(count) 
    Set objFile = objFSO.GetFile(filename) 
    set oShellLink = WshShell.CreateShortcut(strFolder & objFSO.GetBaseName(filename) & ".lnk") 
    oShellLink.TargetPath = filename 
    oShellLink.WindowStyle = 1 
    oShellLink.WorkingDirectory = objFSO.GetParentFolderName(filename) 
    oShellLink.Save 
NEXT 