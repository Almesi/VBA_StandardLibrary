VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "xl_DB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private p_Destination As Variant
Private p_Statements() As Variant
Private p_RecordSet() As Variant
Private p_Record As Variant
Private p_RecordName As String

Private ErrorCatalog(1, 99) As Variant
Private p_Handler As New std_Error

Public Property Let Destination(n_Value As Variant)
    If CheckDestination(n_Value) Then
        If IsObject(n_Value) Then
            Set p_Destination = n_Value
        Else
            p_Destination = n_Value
        End If
    End If
End Property

Public Property Get Destination() As Variant
    If IsObject(p_Destination) Then
        Set Destination = p_Destination
    Else
        Destination = p_Destination
    End If
End Property

Public Property Let Handler(n_Value As std_Error)
    Set p_Handler = n_Value
End Property

Public Property Get Handler() As std_Error
    Set Handler = p_Handler
End Property




Public Function Create(Optional n_Destination As Variant, Optional n_Handler As std_Error) As xl_DB
    Set Create = New xl_DB
    Create.Destination = n_Destination
    Create.Handler = n_Handler
End Function

Private Function CheckDestination(Value As Variant) As Boolean
    Select Case TypeName(Value)
        Case "Workbook", "Worksheet", "Sheet", "Range"
            CheckDestination = True
        Case Else
    End Select
End Function

Public Sub Clear()
    Dim Temp() As Variant
    p_Statements = Temp
End Sub

Public Function Run(ParamArray Statements() As Variant) As Variant
    Dim Temp As Variant
    Temp = Statements
    Run = RunArr(Temp)
End Function

Public Function RunArr(Optional Statements As Variant) As Variant
    Dim i As Long
    Dim CurrentStatement As String
    Dim PreviousValue As Variant

    If Ubound(Statements) <> -1 Then p_Statements = Statements
    For i = 0 To Ubound(p_Statements)
        If CheckStatement(i) Then
            CurrentStatement = p_Statements(i)
            i = i + 1
            PreviousValue = Interpret(i, CurrentStatement, PreviousValue)
        End If
    Next i
    If IsArray(PreviousValue) Then
        If Ubound(PreviousValue, 1) = 0 Then
            If Ubound(PreviousValue, 2) = 0 Then
                RunArr = PreviousValue(0, 0)
                Exit Function
            End If
        End If
    End If
    RunArr = PreviousValue
    
End Function

Public Sub Query(ParamArray Statements() As Variant)
    Dim Temp As Variant
    Temp = Statements
    p_Statements = Temp
End Sub

Public Sub QueryArr(Statements() As Variant)
    p_Statements = Statements
End Sub




Private Function Interpret(Index As Long, CurrentStatement As String, PreviousValue As Variant) As Variant
    Dim Temp As Variant
    Dim i As Long, j As Long, k As Long
    Dim Rng As Range
    Dim Name As String
    On Error GoTo Error
    Select Case CurrentStatement
        Case "SELECT"
            i = Index
            Do Until CheckStatement(i) Or i = Ubound(p_Statements)
                i = i + 1
            Loop
            Interpret = DoSelect(Index, i - 1)
        Case "FROM"
            RecordSet = Index
            Interpret = p_RecordSet
        Case "TOP"
            Interpret = RecordSetPart(PreviousValue, 0, 0, CLng(p_Statements(Index)), Ubound(PreviousValue, 2))
        Case "BOTTOM"
            Interpret = RecordSetPart(PreviousValue, Ubound(PreviousValue, 1) - CLng(p_Statements(Index)), 0, Ubound(PreviousValue, 1), Ubound(PreviousValue, 2))
        Case "AS"
            PreviousValue = p_Statements(Index)
        Case "WHERE"
            Dim Rows() As Long
            Dim Columns() As Long
            Dim Count As Long
            Dim EndIndex As Long
            EndIndex = Index
            Do Until CheckStatement(EndIndex + 1)
                EndIndex = EndIndex + 1
                If EndIndex + 1 > Ubound(p_Statements) Then Exit Do
            Loop
            For i = 0 To Ubound(p_RecordSet, 1)
                If DoWhere(Index, EndIndex, i) Then
                    ReDim Preserve Rows(Count)
                    Rows(Count) = i
                    Count = Count + 1
                End If
            Next i
            ReDim Columns(UBound(PreviousValue, 2))
            For i = 0 To UBound(Columns)
                Columns(i) = GetColumnIndex(PreviousValue(0, i))
            Next i
            If Ubound(Rows)    = -1 Then GoTo Error
            If Ubound(Columns) = -1 Then GoTo Error
            Interpret = RecordSetChopped(p_RecordSet, Rows, Columns)
        Case "CHANGEROWCOUNT"
            Dim TempCount As Long
            Temp = PreviousValue
            ReDim PreviousValue(p_Statements(Index), Ubound(PreviousValue, 2))
            If p_Statements(Index) = -1 Or p_Statements(Index) > UBound(Temp, 1) Then
                For i = 0 To UBound(Temp, 1)
                    For j = 0 To UBound(Temp, 2)
                        PreviousValue(i, j) = Temp(i, j)
                    Next j
                Next i
            Else
                For i = 0 To UBound(Temp, 1)
                    For j = 0 To UBound(Temp, 2)
                        If i = p_Statements(Index) Then
                            i = i + 1
                        End If
                        PreviousValue(TempCount, j) = Temp(i, j)
                    Next j
                    TempCount = TempCount + 1
                Next i
            End If
        Case "DROP"
            Call Clear()
        Case "CHANGECOLUMNCOUNT"
            Dim ColumnIndex As Long
            If IsNumeric(p_Statements(Index)) Then
                i = CLng(p_Statements(Index))
            Else
                For i = 0 To UBound(PreviousValue, 2)
                    If p_Statements(Index) = PreviousValue(0, i) Then Exit For
                Next i
            End If
            ColumnIndex = i
            If ColumnIndex = -1 Or ColumnIndex > UBound(PreviousValue, 2) Then
                ReDim Preserve PreviousValue(UBound(PreviousValue, 1), UBound(PreviousValue, 2) + 1)
            Else
                Temp = PreviousValue
                ReDim Preserve PreviousValue(UBound(PreviousValue, 1), UBound(PreviousValue, 2) - 1)
                For i = 0 To UBound(PreviousValue, 1)
                    For j = 0 To UBound(PreviousValue, 2)
                        If j >= ColumnIndex Then
                            PreviousValue(i, j) = Temp(i, j + 1)
                        Else
                            PreviousValue(i, j) = Temp(i, j)
                        End If
                    Next j
                Next i
            End If
        Case "CHANGEITEM"
            Dim Row As Long
            Row = Clng(p_Statements(Index + 1))
            Select Case TypeName(p_Destination)
                Case "Workbook"
                    Set Rng = p_Destination.Sheets(p_Statements(Index)).Range("A1")
                Case "Worksheet", "Sheet", "Range"
                    Set Rng = FindRange(p_Destination.Rows(1), p_Statements(Index))
                Case Else
            End Select
            j = Index + 2
            Do Until CheckStatement(j)
                If p_Statements(j) = "CURRENTINDEX" Then p_Statements(j) = Row
                Rng.Offset(Row, j - Index - 2).Value = p_Statements(j)
                j = j + 1
                If j > Ubound(p_Statements) Then Exit Do
            Loop
            PreviousValue = Row
        Case "SORT"
            Dim ColOffset As Long
            ColOffset = GetColumnIndex(p_Statements(Index + 1))
            If p_Statements(Index) = True Then
                For i = 0 To UBound(PreviousValue, 1) - 1
                    For j = i To UBound(PreviousValue, 1) - i - 1
                        If PreviousValue(j, ColOffset) > PreviousValue(j + 1, ColOffset) Then
                            For k = 0 To Ubound(PreviousValue, 2)
                                Temp = PreviousValue(j, k)
                                PreviousValue(j, k) = PreviousValue(j + 1, k)
                                PreviousValue(j + 1, k) = Temp
                            Next k
                        End If
                    Next j
                Next i
            Else
                For i = UBound(PreviousValue, 1) To 1 Step -1
                    For j = UBound(PreviousValue, 1) To UBound(PreviousValue, 1) - i + 1 Step -1
                        If PreviousValue(j, ColOffset) > PreviousValue(j - 1, ColOffset) Then
                            For k = 0 To Ubound(PreviousValue, 2)
                                Temp = PreviousValue(j, k)
                                PreviousValue(j, k) = PreviousValue(j - 1, k)
                                PreviousValue(j - 1, k) = Temp
                            Next k
                        End If
                    Next j
                Next i
            End If
            Interpret = PreviousValue
        Case "EXISTS"
            Dim Column As Long
            Dim StartIndex As Long
            Column = GetColumnIndex(p_Statements(Index + 1))
            If Column = -1 Then
                StartIndex = 0
                EndIndex = Ubound(PreviousValue, 2)
            Else
                StartIndex = Column
                EndIndex = Column
            End If
            For i = 0 To UBound(PreviousValue, 1)
                For j = StartIndex To EndIndex
                    If PreviousValue(i, j) = p_Statements(Index) Then
                        Interpret = True
                        Exit Function
                    End If
                Next j
            Next i
        Case "CREATE"
            Select Case TypeName(p_Destination)
                Case "Workbook"
                    Name = p_Statements(Index)
                    Do Until SheetExists(Name) = False
                        Name = p_Statements(Index) & i
                        i = i + 1
                    Loop
                    p_Destination.Sheets.Add
                    ActiveSheet.Name = Name
                    Set Rng = p_Destination.Sheets(Name).Range("A1")
                Case "Worksheet", "Sheet", "Range"
                    p_Destination.Range(p_Statements(Index + 1)).Value = p_Statements(Index)
                    Set Rng = p_Destination.Range(p_Statements(Index + 1)).Offset(1, 0)
                Case Else
            End Select
            For i = 0 To Ubound(PreviousValue, 1)
                For j = 0 To Ubound(PreviousValue, 2)
                    Rng.Offset(i, j) = PreviousValue(i, j)
                Next j
            Next i
            Interpret = PreviousValue
        Case "ADDITEM"
            Select Case TypeName(p_Destination)
                Case "Workbook"
                    Set Rng = p_Destination.Sheets(p_Statements(Index)).Range("A1")
                Case "Worksheet", "Sheet", "Range"
                    Set Rng = FindRange(p_Destination.Rows(1), p_Statements(Index)).Offset(1, 0)
                Case Else
            End Select
            Do While Rng.Offset(i, 0).Formula <> ""
                i = i + 1
            Loop
            j = Index + 1
            Do Until CheckStatement(j)
                If p_Statements(j) = "CURRENTINDEX" Then p_Statements(j) = i
                Rng.Offset(i, j - Index - 1).Value = p_Statements(j)
                j = j + 1
                If j > Ubound(p_Statements) Then Exit Do
            Loop
            PreviousValue = i
        Case "COMMIT"
            Select Case TypeName(p_Record)
                Case "Worksheet", "Sheet"
                    p_Record.Name = p_RecordName
                    For i = 0 To Ubound(p_RecordSet, 1)
                        For j = 0 To Ubound(p_RecordSet, 2)
                            p_Record.Cells(i, j) = ""
                        Next j
                    Next i
                    For i = 0 To Ubound(PreviousValue, 1)
                        For j = 0 To Ubound(PreviousValue, 2)
                            p_Record.Cells(i, j) = PreviousValue(i, j)
                        Next j
                    Next i
                Case "Range"
                    p_Record.Value = p_RecordName
                    For i = 0 To Ubound(p_RecordSet, 1)
                        For j = 0 To Ubound(p_RecordSet, 2)
                            p_Record.Offset(1 + i, j) = ""
                        Next j
                    Next i
                    For i = 0 To Ubound(PreviousValue, 1)
                        For j = 0 To Ubound(PreviousValue, 2)
                            p_Record.Offset(1 + i, j) = PreviousValue(i, j)
                        Next j
                    Next i
            End Select
            Interpret = PreviousValue
        Case "UPDATE"
            Select Case TypeName(p_Destination)
                Case "Workbook"
                    Set Rng = p_Destination.Sheets(p_Statements(Index)).Range("A1")
                Case "Worksheet", "Sheet", "Range"
                    Set Rng = FindRange(p_Destination.Rows(1), p_Statements(Index)).Offset(1, 0)
                Case Else
            End Select
            Range(Rng, Range(Rng.End(xlDown), Rng.End(xlToRight))).Clear
            For i = 0 To Ubound(PreviousValue, 1)
                For j = 0 To Ubound(PreviousValue, 2)
                    Rng.Offset(i, j) = PreviousValue(i, j)
                Next j
            Next i
            Interpret = PreviousValue
        Case "PRINT"
            Dim Message As String
            For i = 0 To Ubound(PreviousValue, 1)
                For j = 0 To UBound(PreviousValue, 2)
                    Message = Message & " | " & PreviousValue(i, j)
                Next j
                Debug.Print Message
                Message = ""
            Next i
        Case Else
            Interpret = PreviousValue
    End Select
    Exit Function

    Error:
    Interpret = Empty
End Function

Private Function CheckStatement(Index As Long) As Boolean
    Select Case UCase(p_Statements(Index))
        Case "SELECT"
        Case "FROM"
        Case "TOP"
        Case "BOTTOM"
        Case "AS"
        Case "WHERE"
        Case "CHANGEROWCOUNT"
        Case "DROP"
        Case "CHANGECOLUMNCOUNT"
        Case "CHANGEITEM"
        Case "SORT"
        Case "EXISTS"
        Case "CREATE"
        Case "ADDITEM"
        Case "COMMIT"
        Case "UPDATE"
        Case "PRINT"
        Case Else
            Exit Function
    End Select
    CheckStatement = True
End Function

Private Property Let RecordSet(Index As Long)
    Dim Temp as Range
    Dim i As Long, j As Long
    On Error GoTo Error
    Select Case TypeName(p_Destination)
        Case "Workbook"
            Set Temp = p_Destination.Sheets(p_Statements(Index)).Range("A1")
            Set p_Record = p_Destination.Sheets(p_Statements(Index))
            p_RecordName = p_Record.Name
        Case "Worksheet", "Sheet", "Range"
            If p_Destination.Range("A1").Value = p_Statements(Index) Then ' If the first Cell is the searched database, as .find wont look into the first cell
                Set Temp = p_Destination.Range("A2")
            Else
                Set Temp = FindRange(p_Destination.Rows(1), p_Statements(Index)).Offset(1, 0)
            End If
            Set p_Record = Temp.Offset(-1, 0)
            p_RecordName = p_Record.Value
        Case Else
    End Select
    Set Temp = Range(Temp, Range(Temp.End(xlDown), Temp.End(xlToRight)))

    ReDim p_RecordSet(Temp.Rows.Count - 1, Temp.Columns.Count - 1)
    For i = 0 To UBound(p_RecordSet, 1)
        For j = 0 To UBound(p_RecordSet, 2)
            p_RecordSet(i, j) = Temp.Cells(i + 1, j + 1)
        Next j
    Next i
    Error:
End Property

Private Property Get RecordSetItem(Row As Long, ColumnNameOrIndex As Variant) As Variant
    Dim i As Long
    If IsNumeric(ColumnNameOrIndex) Then
        ColumnNameOrIndex = CLng(ColumnNameOrIndex)
    Else
        For i = 0 To Ubound(p_RecordSet, 2)
            If p_RecordSet(0, i) = ColumnNameOrIndex Then
                ColumnNameOrIndex = i
                Exit For
            End If
        Next i
    End If
    If IsNumeric(ColumnNameOrIndex) = False Then Exit Property
    RecordSetItem = p_RecordSet(Row, ColumnNameOrIndex)
End Property

Private Function RecordSetPart(Arr As Variant, StartRow As Long, StartColumn As Long, EndRow As Long, EndColumn As Long) As Variant()
    Dim i As Long, j As Long
    Dim Temp() As Variant
    ReDim Temp(EndRow - StartRow + 1, EndColumn - StartColumn + 1)
    For i = StartRow To EndRow
        For j = StartColumn To EndColumn
            Temp(i - StartRow, j - StartColumn) = Arr(i, j)
        Next j
    Next i
    RecordSetPart = Temp
End Function

Private Function RecordSetChopped(Arr As Variant, Rows() As Long, Columns() As Long) As Variant()
    Dim i As Long, j As Long
    Dim Temp() As Variant
    ReDim Temp(Ubound(Rows), UBound(Columns))
    For i = 0 To Ubound(Rows)
        For j = 0 To UBound(Columns)
            Temp(i, j) = Arr(Rows(i), Columns(j))
        Next j
    Next i
    RecordSetChopped = Temp
End Function

Private Function DoWhere(StartIndex As Long, EndIndex As Long, CurrentIndex As Long) As Boolean
    Dim i As Long, j As Long
    For i = StartIndex To EndIndex
        Select Case p_Statements(i)
            Case "="
                DoWhere = Iif(GetCurrent(p_Statements(i - 1), CurrentIndex) = GetCurrent(p_Statements(i + 1), CurrentIndex), True, False)
                i = i + 1
            Case ">"
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) > GetCurrent(p_Statements(i + 1), CurrentIndex)
                i = i + 1
            Case "<"
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) < GetCurrent(p_Statements(i + 1), CurrentIndex)
                i = i + 1
            Case ">=", "=>"
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) >= GetCurrent(p_Statements(i + 1), CurrentIndex)
                i = i + 1
            Case "=<", "<="
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) =< GetCurrent(p_Statements(i + 1), CurrentIndex)
                i = i + 1
            Case "<>", "NOT"
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) <> GetCurrent(p_Statements(i + 1), CurrentIndex)
                i = i + 1
            Case "LIKE"
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) Like GetCurrent(p_Statements(i + 1), CurrentIndex)
                i = i + 1
            Case "BETWEEN"
                DoWhere = GetCurrent(p_Statements(i - 1), CurrentIndex) > GetCurrent(p_Statements(i + 1), CurrentIndex) And GetCurrent(p_Statements(i - 1), CurrentIndex) < GetCurrent(p_Statements(i + 2), CurrentIndex)
                i = i + 2
            Case "&&", "AND"
                DoWhere = DoWhere(i - 2, i - 2, CurrentIndex) And DoWhere(i + 2, i + 2, CurrentIndex)
                i = i + 3
            Case "||", "OR"
                DoWhere = DoWhere(i - 2, i - 2, CurrentIndex) Or DoWhere(i + 2, i + 2, CurrentIndex)
                i = i + 3
            Case "IS EMPTY"
                DoWhere = IsEmpty(p_Statements(i + 1))
                i = i + 1
            Case "IN"
                For j = i + 1 To EndIndex
                    If GetCurrent(p_Statements(i - 1), CurrentIndex) = p_Statements(j) Then
                        DoWhere = True
                        Exit For
                    End If
                Next
            Case Else
        End Select
    Next i
End Function

Private Function DoSelect(StartIndex As Long, EndIndex As Long) As Variant()
    Dim i As Long, j As Long, k As Long
    Dim Temp() As Variant
    If p_Statements(StartIndex) = "*" Then
        Temp = p_RecordSet
    Else
        If EndIndex - StartIndex + 1 = 1 Then
            ReDim Temp(Ubound(p_RecordSet, 1), 0)
        Else
            ReDim Temp(Ubound(p_RecordSet, 1), EndIndex - StartIndex)
        End If
        For i = StartIndex To EndIndex
            For j = 0 To UBound(p_RecordSet, 2)
                If p_RecordSet(0, j) = p_Statements(i) Then
                    For k = 0 To Ubound(p_RecordSet, 1)
                        Temp(k, i - StartIndex) = p_RecordSet(k, j)
                    Next k
                End If
            Next j
        Next i
    End If
    DoSelect = Temp
End Function

Private Function GetCurrent(Value As Variant, CurrentRow As Long) As Variant
    Select Case True
        Case UCase(Value) Like "CURRENTROWCOLUMN:*"
            GetCurrent = RecordSetItem(CurrentRow, Mid(Value, InStr(1, Value, ":") + 1, Len(Value)))
        Case Else
            GetCurrent = Value
    End Select
End Function

Private Function GetColumnIndex(Value As Variant) As Long
    Dim i As Long
    If IsNumeric(Value) Then
        GetColumnIndex = CLng(Value)
        Exit Function
    Else
        For i = 0 To UBound(p_RecordSet, 2)
            If p_RecordSet(0, i) = Value Then
                GetColumnIndex = i
                Exit Function
            End If
        Next i
    End If
    GetColumnIndex = -1
End Function

Private Function SheetExists(Name As String) As Boolean
    Dim ws As Worksheet
    For Each ws In p_Destination.Worksheets
        If ws.Name = Name Then
            SheetExists = True
            Exit Function
        End If
    Next ws
End Function

Private Function FindRange(Rng As Range, Value As Variant) As Range
    If Rng.Cells(1, 1).Value = Value Then
        Set FindRange = Rng.Cells(1, 1)
    Else
        Set FindRange = Rng.Find(Value,,,xlWhole,,,True)
    End If
End Function