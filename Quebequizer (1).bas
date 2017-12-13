Attribute VB_Name = "Quebequizer"
Private Sub PT_Quebequizer()
' Convert_formatting_to_QC Macro
Dim vSelectedRange As Range
Dim vCell As Range
On Error GoTo exitmacro
Set vSelectedRange = Application.InputBox("Veuillez choisir une plage de cellules à convertir.", "Quebequizer", , , , , , 8)


For Each vCell In vSelectedRange
On Error GoTo nextvCell
    If IsEmpty(vCell) Then
        GoTo nextvCell
    ElseIf Not IsNumeric(vCell.Value) Then
        vCell.Replace what:=" ", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        vCell.Replace what:=",", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        vCell.Replace what:=".", Replacement:=",", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        vCell.Replace what:="$", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        vCell.Replace what:="€", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        vCell.Value = CDec(vCell.Value)
        vCell.Style = "Currency"
    End If
    
nextvCell:
Next vCell
Call MsgBox("Conversion terminée!", vbOKOnly + vbInformation, "Quebequizer")
exitmacro:
End Sub



Private Sub cellsAutoFit()

'This macro autofits all cells in the current sheet.

ActiveSheet.Cells.Columns.AutoFit
ActiveSheet.Cells.Rows.AutoFit

End Sub



