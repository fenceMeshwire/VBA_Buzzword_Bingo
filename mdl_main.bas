Option Explicit

Sub create_bingo()

Dim functions As New cls_functions

Dim intRow, intColumn, intRandom As Integer
Dim varWords As Variant
Dim wksMaterial, wksSheet As Worksheet

Set wksMaterial = wordlist
Set wksSheet = bingo

wksSheet.Visible = xlSheetVisible

CallByName functions, "clear_area", VbMethod, wksSheet
CallByName functions, "scale_print", VbMethod, wksSheet
CallByName functions, "maintain_cell_dimensions", VbMethod, wksSheet
varWords = CallByName(functions, "get_buzzwords", VbMethod, wksSheet)

For intRow = 1 To 10
  For intColumn = 1 To 10
    intRandom = CallByName(functions, "create_random_number", VbMethod)
    wksSheet.Cells(intRow, intColumn).Value = wksMaterial.Cells(intRandom, 1).Value
  Next intColumn
Next intRow

CallByName functions, "frame_area", VbMethod, wksSheet
CallByName functions, "copy_save_close", VbMethod, wksSheet

wksSheet.Visible = xlSheetVeryHidden

End Sub
