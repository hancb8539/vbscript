Function CreateGUID
  Dim TypeLib
  Set TypeLib = CreateObject("Scriptlet.TypeLib")
  CreateGUID = Mid(TypeLib.Guid,2,36)
End Function

tmp = "{"+CreateGUID()+"}" '隨機生成GUID

WScript.Echo tmp



Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
 
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")
 
strKeyPath = "SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths\"+tmp '黑名單路徑
oReg.CreateKey HKEY_LOCAL_MACHINE,strKeyPath

strname = "ItemData"
strvalue = "C:\Program Files\Adobe\Reader 11.0\Reader\"
oReg.SetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strname,strvalue

dwname = "SaferFlags"
dwValue = 0
oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,dwname,dwValue