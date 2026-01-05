#Requires AutoHotkey v2.0

; class ListViewHistory
;
; A class to manage cut/copy/paste operations in a ListView.
; The rows of the ListView can have checkboxes and icons.
;
; Author: iPhilip
;
; The class is instantiated with the following parameters:
;   LV - A Guicontrol object representing a ListView
;   IconNumbers - (Optional) An array of icon indices representing the icons associated with the rows of the ListView. If omitted, it defaults to an empty array.
;   MaxHistoryEntries - (Optional) An integer representing the maximum number of history entries. If omitted, it defaults to 0, i.e. no limit.
;
; User Properties:
;
; NumberOfEntries
;
; User Methods:
;
; Cut(RowType?)
; Copy(RowType?)
; Paste(Below := false)
; Undo()
; Clear()

class ListViewHistory
{
   __New(LV, IconNumbers?, MaxHistoryEntries?) {
      this.LV := LV
      this.History := []
      this.IconNumbers := IconNumbers ?? []
      this.MaxHistoryEntries := MaxHistoryEntries ?? 0
   }
   
   ; ---------------
   ; User properties
   ; ---------------
   
   NumberOfEntries => this.History.Length
   
   ; ------------
   ; User methods
   ; ------------
   
   ; Cuts the row types and saves the information about the rows in history.
   ; If RowType is blank or omitted, the method cuts the selected/highlighted rows.
   ; If RowType is 'C' or 'Checked', the method cuts the checked rows.
   ; If RowType is 'F' or 'Focused', the method cuts the focused row.
   ; The method eturns the number of rows cut.
   
   Cut(RowType?) {
      Rows := Map()  ; Map of arrays representing the selected rows.
      Rows.Event := 'Cut'  ; Identify the event. This is used in the Undo method.
      ;
      ; Identify and save the rows to be cut.
      ;
      RowsToCut := []
      RowNumber := 0
      Loop {
         RowNumber := this.LV.GetNext(RowNumber, RowType?)
         if not RowNumber
            break
         Rows[RowNumber] := this.CopyRow(RowNumber)
         RowsToCut.InsertAt(1, RowNumber)
      }
      if not Rows.Count
         return 0
      ;
      ; Delete the rows from the bottom to the top.
      ;
      for RowNumber in RowsToCut
         this.LV.Delete(RowNumber)
      ;
      ; Focus and select the last cut row.
      ;
      RowNumber := RowsToCut[RowsToCut.Length]
      NumberOfRows := this.LV.GetCount()
      this.LV.Modify(RowNumber > NumberOfRows ? NumberOfRows : RowNumber, 'Focus Select')
      ;
      ; Add the cut rows to history.
      ;
      this.History.Push(Rows)
      if this.MaxHistoryEntries && this.History.Length > this.MaxHistoryEntries
         this.History.RemoveAt(1)
      ;
      ; Return the number of cut rows.
      ;
      return Rows.Count
   }
   
   ; Copies the row types and saves the information about the rows in history.
   ; If RowType is blank or omitted, the method copies the selected/highlighted rows.
   ; If RowType is 'C' or 'Checked', the method copies the checked rows.
   ; If RowType is 'F' or 'Focused', the method copies the focused row.
   ; The method eturns the number of rows copied.
   
   Copy(RowType?) {
      Rows := Map()  ; Map of arrays representing the selected rows.
      Rows.Event := 'Copy'  ; Identify the event. This is used in the Undo method.
      ;
      ; Identify and save the rows to be copied.
      ;
      RowNumber := 0
      Loop {
         RowNumber := this.LV.GetNext(RowNumber, RowType?)
         if not RowNumber
            break
         Rows[RowNumber] := this.CopyRow(RowNumber)
      }
      if not Rows.Count
         return 0
      ;
      ; Add the copied rows to history.
      ;
      this.History.Push(Rows)
      if this.MaxHistoryEntries && this.History.Length > this.MaxHistoryEntries
         this.History.RemoveAt(1)
      ;
      ; Return the number of copied rows.
      ;
      return Rows.Count
   }
   
   ; Paste the most recent history into the ListView.
   ; If the Below parameter is true, the most recent history is pasted below the bottom-most selected row.
   ; If no rows are selected, the most recent histoy is pasted at the bottom of the ListView.
   ; If the Below parameter is false, the most recent history is pasted above the top-most selected row.
   ; If no rows are selected, the most recent history is pasted at the top of the ListView.
   ; Returns the number of rows pasted into the ListView.
   
   Paste(Below := false) {
      if not this.History.Length
         return 0
      Rows := this.History.Pop()
      RowNumbers := []
      if Below {
         ;
         ; Get the bottom-most selected row.
         ;
         RowNumber := 0
         Loop {
            SelectedRowNumber := this.LV.GetNext(RowNumber)
            if not SelectedRowNumber
               break
            RowNumber := SelectedRowNumber
         }
         ;
         ; If there was a bottom-most selected row, insert the saved rows below that one.
         ; Otherwise, add the saved rows to the bottom.
         ;
         if RowNumber {
            for , Row in Rows
               this.LV.Insert(++RowNumber, (Row.CheckState = 1 ? 'Check ' : '') 'Icon' Row.IconNumber, Row*), RowNumbers.Push(RowNumber)
         } else {
            for , Row in Rows
               RowNumbers.Push(this.LV.Add((Row.CheckState = 1 ? 'Check ' : '') 'Icon' Row.IconNumber, Row*))
         }
      } else {
         ;
         ; Get the first selected row.
         ;
         RowNumber := this.LV.GetNext()
         ;
         ; If a row was selected, insert the saved rows above that one.
         ; Otherwise, insert the selected rows at the top.
         ;
         if RowNumber {
            for , Row in Rows
               this.LV.Insert(RowNumber, (Row.CheckState = 1 ? 'Check ' : '') 'Icon' Row.IconNumber, Row*), RowNumbers.Push(RowNumber++)
         } else {
            for , Row in Rows
               RowNumbers.Push(this.LV.Insert(A_Index, (Row.CheckState = 1 ? 'Check ' : '') 'Icon' Row.IconNumber, Row*))
         }
      }
      ;
      ; Select the pasted rows.
      ;
      this.LV.Modify(0, '-Focus -Select')
      for RowNumber in RowNumbers
         this.LV.Modify(RowNumber, A_Index = 1 ? 'Focus Select' : 'Select')
      ;
      ; Return the number of pasted rows.
      ;
      return Rows.Count
   }
   
   ; Undoes the most recent cut/copy operation.
   ; Undoing a copy operation does not visually affect the ListView.
   ; Other operations, e.g. Paste, Clear, cannot be undone.
   ; Returns the number of rows removed from history.
   
   Undo() {
      if not this.History.Length
         return 0
      Rows := this.History.Pop()
      if Rows.Event = 'Cut'
         for RowNumber, Row in Rows
            this.LV.Insert(RowNumber, (Row.CheckState = 1 ? 'Check ' : '') 'Icon' Row.IconNumber, Row*)
      return Rows.Count
   }
   
   ; Clears the history.
   ; Returns the previous number of history entries.
   
   Clear() {
      Length := this.History.Length
      this.History.Length := 0
      return Length
   }
   
   ; --------------
   ; Helper methods
   ; --------------
   
   ; Gets information about the row: contents, checkstate, and icon number (if any).
   
   CopyRow(RowNumber) {
      Row := []
      Loop this.LV.GetCount('Column')
         Row.Push(this.LV.GetText(RowNumber, A_Index))
      Row.CheckState := this.GetCheckState(RowNumber)
      Row.IconNumber := this.IconNumbers.Has(RowNumber) ? this.IconNumbers[RowNumber] : -1
      return Row
   }
   
   ; Gets the state of the checkbox: -1 for hidden, 0 for unchecked, and 1 for checked.
   
   GetCheckState(Row) {
      static LVM_GETITEMSTATE := 0x102C
      static LVIS_STATEIMAGEMASK := 0xF000
      
      return (SendMessage(LVM_GETITEMSTATE, Row - 1, LVIS_STATEIMAGEMASK, this.LV) >> 12) - 1
   }
}
