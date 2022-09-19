Option Explicit

Public Function scale_print(wksSheet As Worksheet)

With wksSheet
    .PageSetup.PaperSize = xlPaperA4
    .PageSetup.Orientation = xlLandscape
    .PageSetup.Zoom = False
    .PageSetup.FitToPagesWide = 1
    .PageSetup.FitToPagesTall = 1
End With

End Function

Public Function maintain_cell_dimensions(wksSheet As Worksheet)

Dim intRow, intColumn As Integer
Dim rngBingoArea As Range

Set rngBingoArea = wksSheet.Range("A1:E5")

rngBingoArea.RowHeight = 100
rngBingoArea.ColumnWidth = 30

End Function

Public Function create_random_number() As Integer

Dim intCounter, intRandom As Integer
Dim intCellFree As Integer

intCellFree = wordlist.Cells(wordlist.Rows.Count, 1).End(xlUp).Row

intRandom = Int((intCellFree * Rnd) + 1)    ' Generate random value between 1 and 10.

create_random_number = intRandom

End Function

Public Function get_buzzwords(wksMaterial As Worksheet) As Variant

Dim intCounter, intRow, intRowMax As Integer
Dim varWords As Variant

intCounter = 0

intRowMax = wordlist.Cells(wordlist.Rows.Count, 1).End(xlUp).Row
ReDim varWords(intCounter)
For intRow = 1 To intRowMax
  varWords(intCounter) = wordlist.Cells(intRow, 1).Value
  intCounter = intCounter + 1
  ReDim Preserve varWords(intCounter)
Next intRow

ReDim Preserve varWords(UBound(varWords) - 1)

get_buzzwords = varWords

End Function

Public Function clear_area(wksSheet As Worksheet)

wksSheet.UsedRange.Clear

End Function

Public Function frame_area(wksSheet As Worksheet)

Dim intRow, intColumn As Integer
Dim rngFrame As Range

Set rngFrame = wksSheet.Range(wksSheet.Cells(1, 1), wksSheet.Cells(5, 5))

With rngFrame
  .Borders(xlInsideHorizontal).LineStyle = xlContinuous
  .Borders(xlInsideVertical).LineStyle = xlContinuous
  .BorderAround Weight:=xlThick, ColorIndex:=1
  .HorizontalAlignment = xlCenter
  .VerticalAlignment = xlCenter
End With

End Function

Public Function copy_save_close(wksSheet As Worksheet)

wksSheet.Copy
ActiveWorkbook.SaveAs (ThisWorkbook.Path & "\Bingo.xlsx")
ActiveWorkbook.Close

End Function
