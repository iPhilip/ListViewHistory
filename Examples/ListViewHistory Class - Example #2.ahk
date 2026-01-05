; Elaborate demonstration of the ListViewHistory class methods applied to three different ListView controls.

#Requires AutoHotkey v2.0
#Include ListViewHistory Class.ahk

LV_Parameters := ['w350 r10', ['Col 1', 'Col 2', 'Col 3']]
LV_RowContent(LV, Index) => ['LV ' LV ' Row ' Index ' Col 1', 'LV ' LV ' Row ' Index ' Col 2', 'LV ' LV ' Row ' Index ' Col 3']

MyGui := Gui()
MyGui.OnEvent('Close', (*) => ExitApp())

MyGui.AddText( , 'Regular ListView')
LV1 := MyGui.AddListView(LV_Parameters*)
Loop 10
   LV1.Add( , LV_RowContent(1, A_Index)*)
LV1.ModifyCol()
LV1.OnEvent('Click', LV1_Click)

LV1_Click(*) {
   global
   if History != History1 {
      History := History1
      RowType := UnSet
   }
}

MyGui.AddText( , 'ListView with checkboxes (Select rows using checkboxes)')
LV2 := MyGui.AddListView('Checked ' LV_Parameters[1], LV_Parameters[2])
Loop 10
   LV2.Add( , LV_RowContent(2, A_Index)*)
LV2.ModifyCol()
LV2.OnEvent('Click', LV2_Click)
LV2.OnNotify(LVN_ITEMCHANGING := -100, OnItemChanging)

LV2_Click(LV, Info) {
   global
   if History != History2 {
      History := History2
      RowType := 'Checked'
   }
}

MyGui.AddText( , 'ListView with checkboxes and icons (Select rows using checkboxes)')
LV3 := MyGui.AddListView('Checked ' LV_Parameters[1], LV_Parameters[2])
ImageListID := IL_Create(10)
LV3.SetImageList(ImageListID)
Loop 10
   IL_Add(ImageListID, 'shell32.dll', A_Index) 

Icons := []
Loop 10
   LV3.Add('Icon' A_Index, LV_RowContent(3, A_Index)*), Icons.Push(A_Index)
LV3.ModifyCol()
LV3.OnEvent('Click', LV3_Click)
LV3.OnNotify(LVN_ITEMCHANGING, OnItemChanging)

LV3_Click(LV, Info) {
   global
   if History != History3 {
      History := History3
      RowType := 'Checked'
   }
}

OnItemChanging(LV, lParam) {
   Critical
   lParam += 3 * A_PtrSize
   static LVIS_STATEIMAGEMASK := 0xF000
   RowNumber := NumGet(lParam, 0, 'Int') + 1
   NewState  := NumGet(lParam, 8, 'UInt') & LVIS_STATEIMAGEMASK
   if NewState {
      LV.OnNotify(LVN_ITEMCHANGING, OnItemChanging, 0)
      LV.Modify(RowNumber, NewState >> 12 > 1 ? '+Focus +Select' : '-Focus -Select')
      LV.OnNotify(LVN_ITEMCHANGING, OnItemChanging)
      return false
   }
   return true
}

MyGui.Show()

History := History1 := ListViewHistory(LV1)
History2 := ListViewHistory(LV2)
History3 := ListViewHistory(LV3, Icons)

#HotIf WinActive(MyGui)

^x::MsgBox History.Cut(RowType?)
^c::MsgBox History.Copy(RowType?)
^v::MsgBox History.Paste()
^b::MsgBox History.Paste(true)
^u::MsgBox History.Undo()

F5::MsgBox History.Clear()

F6::History2.History.Push(History1.History.Pop())  ; Move the most recent history from ListView 1 to ListView 2.
F7::History2.History.Push(History3.History.Pop())  ; Move the most recent history from ListView 3 to ListView 2.

; Note: With respect to icons, moving history from one ListView to another, only transfers the icon number - not the image itself.

#HotIf

/*
typedef struct tagNMLISTVIEW {
  NMHDR  hdr;           3 * A_PtrSize
  int    iItem;         4
  int    iSubItem;      4
  UINT   uNewState;     4
  UINT   uOldState;     4
  UINT   uChanged;      4
  POINT  ptAction;      4 + A_PtrSize
  LPARAM lParam;        A_PtrSize
} NMLISTVIEW            24 + 5 * A_PtrSize