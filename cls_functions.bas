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

Public Function check_result_vertical()

Dim bleBingoVertical As Boolean
Dim intColumn, intRow As Integer
Dim intCounter As Integer

Dim wksSheet As Worksheet
Set wksSheet = bingo

' Vertical Bingo
bleBingoVertical = False
For intColumn = 1 To 5
  For intRow = 1 To 5
    If bingo.Cells(intRow, intColumn).Interior.ColorIndex = 4 Then
      bleBingoVertical = True
      intCounter = intCounter + 1
      If intCounter = 5 Then
        check_result_vertical = bleBingoVertical
        Exit Function
      Else
        bleBingoVertical = False
      End If
    End If
  Next intRow
  intCounter = 0
Next intColumn

End Function

Public Function check_result_horizontal()

Dim bleBingoHorizontal As Boolean
Dim intColumn, intRow As Integer
Dim intCounter As Integer

Dim wksSheet As Worksheet
Set wksSheet = bingo

' Horizontal Bingo
For intRow = 1 To 5
  For intColumn = 1 To 5
    If bingo.Cells(intRow, intColumn).Interior.ColorIndex = 4 Then
      bleBingoHorizontal = True
      intCounter = intCounter + 1
      If intCounter = 5 Then
        check_result_horizontal = bleBingoHorizontal
        Exit Function
      End If
    Else
      bleBingoHorizontal = False
    End If
  Next intColumn
  intCounter = 0
Next intRow

End Function

Public Function check_result_diagonal()

Dim bleBingoDiagonal As Boolean
Dim intColumn, intRow As Integer

Dim wksSheet As Worksheet
Set wksSheet = bingo

' Diagonal Bingo top left to bottom right
If bingo.Cells(1, 1).Interior.ColorIndex = 4 Then
  If bingo.Cells(2, 2).Interior.ColorIndex = 4 Then
    If bingo.Cells(3, 3).Interior.ColorIndex = 4 Then
      If bingo.Cells(4, 4).Interior.ColorIndex = 4 Then
        If bingo.Cells(5, 5).Interior.ColorIndex = 4 Then
          bleBingoDiagonal = True
          check_result_diagonal = bleBingoDiagonal
          Exit Function
        End If
      End If
    End If
  End If
End If

' Diagonal Bingo bottom left to top right
If bingo.Cells(1, 5).Interior.ColorIndex = 4 Then
  If bingo.Cells(2, 4).Interior.ColorIndex = 4 Then
    If bingo.Cells(3, 3).Interior.ColorIndex = 4 Then
      If bingo.Cells(4, 2).Interior.ColorIndex = 4 Then
        If bingo.Cells(5, 1).Interior.ColorIndex = 4 Then
          bleBingoDiagonal = True
          check_result_diagonal = bleBingoDiagonal
          Exit Function
        End If
      End If
    End If
  End If
End If

check_result_diagonal = False

End Function

Public Function check_bingo_result()

Dim bleResult As Boolean
Dim intColumn, intRow As Integer

Dim wksSheet As Worksheet
Set wksSheet = bingo

'If check_result_horizontal = True Then bleResult = True
'If check_result_vertical = True Then bleResult = True
'If check_result_diagonal = True Then bleResult = True
 
If check_result_horizontal = True Or check_result_vertical = True Or check_result_diagonal = True Then
  MsgBox "BINGO!!!"
  If MsgBox("Do you want to restart the game?", vbYesNo) = vbYes Then
    Call create_bingo
  End If
End If

End Function
