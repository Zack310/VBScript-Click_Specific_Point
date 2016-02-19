'Click after imputed number'
Dim clk, obj
set obj=createobject("wscript.shell")

'x 550 750 950
'y 710 580 450


Do 
clk=inputbox("input here","Click Number",vbSystemModal)
if clk=vbCancel Then

ElseIf clk=1 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 550, 710"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"

	
Elseif clk=2 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 750, 710"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"

Elseif clk=3 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 950, 710"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"

Elseif clk=4 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 550, 580"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"


Elseif clk=5 Then '元の場所は真ん中の５'
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 750, 580"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"
	
Elseif clk=6 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 950, 580"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"
	
Elseif clk=7 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 550, 450"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"
	
Elseif clk=8 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 750, 450"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"

Elseif clk=9 Then
	obj.run "rundll32.exe MouseEmulator.dll, _SetMouseXY@16 950, 450"
	obj.run "rundll32.exe MouseEmulator.dll, _ClickLeft@16"
	wscript.sleep 50
	obj.AppActivate "Click Number"

End if
Loop Until clk=999

