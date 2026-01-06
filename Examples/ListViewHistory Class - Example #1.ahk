; Simple demonstration of the ListViewHistory class Cut and Paste methods.

#Requires AutoHotkey v2.0
#Include ..\Lib\ListViewHistory Class.ahk

LV_Parameters := ['w275 r10', ['Col 1', 'Col 2', 'Col 3']]
LV_RowContent(Index) => ['Row ' Index ' Col 1', 'Row ' Index ' Col 2', 'Row ' Index ' Col 3']

MyGui := Gui()
MyGui.OnEvent('Close', (*) => ExitApp())

MyGui.AddText( , 'Regular ListView')
LV := MyGui.AddListView(LV_Parameters*)
Loop 100
   LV.Add( , LV_RowContent(A_Index)*)
LV.ModifyCol()

MyGui.Show()

History := ListViewHistory(LV)
Loop 50
   LV.Modify(A_Index + 50, 'Select')
History.Cut()

MsgBox

LV.Modify(0, '-Select')
LV.Modify(1, 'Select')

History.Paste()
