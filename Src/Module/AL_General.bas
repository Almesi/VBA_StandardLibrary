Attribute VB_Name = "AL_General"

' Swap any 2 Variables
Sub AL_Swap(ByRef Variable1 As Variant, ByRef Variable2 As Variant)

    Dim Temp As Variant
    Temp = Variable1
    Variable1 = Variable2
    Variable2 = Temp

End Sub

Sub AL_General_FlipRange(InputRange As Range, Optional FlipDirection As Boolean = False)

    Dim FlippedRange As Range
    Dim LastColumn As Long
    Dim LastRow As Long
    Dim i As Long, j As Long

    Set FlippedRange = Range(ActiveCell, ActiveCell.Offset(InputRange.Rows.Count, InputRange.Columns.Count))
    LastColumn = InputRange.Columns.Count
    LastRow = InputRange.Rows.Count
    If FlipDirection = True Then
            For i = 0 To LastRow
                For j = 0 To LastColumn
                    FlippedRange.Cells(i, j + 1).Formula = InputRange.Cells(i, LastColumn - j).Formula
                Next
            Next
        Else
            For i = 0 To LastColumn
                For j = 0 To LastRow
                    FlippedRange.Cells(j, i + 1).Formula = InputRange.Cells(LastRow + 1 - j, i + 1).Formula
                Next
            Next
    End If
    
End Sub

Sub AL_General_TransposeRange(InputRange As Range)
    
    Dim TransposedRange As Range
    Dim i As Long
    Dim j As Long

    Set TransposedRange = Range(ActiveCell, ActiveCell.Offset(InputRange.Columns.Count - InputRange.Column, InputRange.Rows.Count - InputRange.Row))
    For i = 1 To InputRange.Rows.Count
        For j = 1 To InputRange.Columns.Count
            TransposedRange.Cells(j, i).Formula = InputRange.Cells(i, j).Formula
        Next j
    Next i

End Sub