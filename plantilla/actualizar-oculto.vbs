' Corre el actualizador sin mostrar ventana.
Set sh = CreateObject("WScript.Shell")
d = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & d & "\Actualizar-Addins-VENTAS.ps1""", 0, False
