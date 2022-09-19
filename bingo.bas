Option Explicit

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, _
 Cancel As Boolean)

Cancel = True

Dim rngCell, rngTargetCells As Range

Set rngTargetCells = Range("A1:E5")

For Each rngCell In Target.Cells
  If rngCell.Column < 6 Then
    If rngCell.Row < 6 Then
      rngCell.Interior.ColorIndex = 4
    End If
  End If
Next rngCell

End Sub
