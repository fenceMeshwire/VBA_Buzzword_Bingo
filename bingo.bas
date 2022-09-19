Option Explicit

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, _
 Cancel As Boolean)

Cancel = True

Dim rngCell, rngTargetCells As Range

Set rngTargetCells = Range("A1:E5")

For Each rngCell In Target.Cells
  rngCell.Interior.ColorIndex = 4
  Debug.Print rngCell
Next rngCell

End Sub
