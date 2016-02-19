'Mouse Test'
Dim Obj
Set Obj=CreateObject("wscript.shell")
wscript.sleep 2000
obj.run "rundll32.exe MouseEmulator.dll, _RotWheel@16 120"
