Set WMPlayer = CreateObject("WMPlayer.OCX.7")
WMPlayer.CdRomCollection.Item(0).Eject()
MsgBox "Text", 64, "Caption"
'WScript.Sleep(2000)
WMPlayer.CdRomCollection.Item(0).Eject()