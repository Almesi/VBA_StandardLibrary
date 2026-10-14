VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private mFieldDefinitions As Collection
Private mDataEntries()    As DataEntry
Private mFieldCount       As Long
Private mRowCount         As Long


Private Sub Class_Initialize()
    Call Clear()
End Sub

Public Sub Clear()
    Set mFieldDefinitions = New Collection

    Erase mDataEntries

    mFieldCount = 0
    mRowCount = 0
End Sub


'============================================================
'====================[Getters/Setters]=======================
'============================================================

Public Property Get FieldDefinitions() As Collection
    Set FieldDefinitions = mFieldDefinitions
End Property


Public Function DataEntries() As DataEntry()
    DataEntries = mDataEntries
End Property


Public Property Get FieldCount() As Long
    FieldCount = mFieldCount
End Property


Public Property Get RowCount() As Long
    RowCount = mRowCount
End Property


Public Property Get FieldDefinition(ByVal Index As Long) As FieldDefinition
    If Index < 1 Or Index > mFieldCount Then Exit Property

    Set FieldDefinition = mFieldDefinitions(Index)
End Property


Public Property Get DataEntryAt(ByVal RowIndex As Long, ByVal FieldIndex As Long) As DataEntry
    If RowIndex   < 1 Or RowIndex   > mRowCount   Then Exit Property
    If FieldIndex < 1 Or FieldIndex > mFieldCount Then Exit Property

    Set DataEntryAt = mDataEntries(RowIndex, FieldIndex)
End Property

Public Property Get FieldEntryAt(ByVal RowIndex As Long, ByVal FieldIndex As Long) As FieldEntry
    If RowIndex   < 1 Or RowIndex   > mRowCount   Then Exit Property
    If FieldIndex < 1 Or FieldIndex > mFieldCount Then Exit Property

    Set FieldEntryAt = FieldEntry.Create(mFieldDefinitions(FieldIndex), DataEntryAt(RowIndex, FieldIndex))
End Property

Public Function FieldByName(ByVal Name As String) As Field
    Dim Index As Long

    For Index = 1 To mFieldCount
        If StrComp(mFieldDefinitions(Index).Name, Name, vbTextCompare) = 0 Then
            Set FieldByName = Field(Index)
            Exit Function
        End If
    Next Index
    Set FieldByName = Nothing
End Function


Public Function Field(ByVal Index As Long) As Field
    Dim Result   As Field
    Dim RowIndex As Long
    Dim Entry    As DataEntry

    If Index < 1 Or Index > mFieldCount Then Exit Function

    Set Result = New Field
    Set Result.Definition = mFieldDefinitions(Index)

    For RowIndex = 1 To mRowCount
        Set Entry = DataEntryAt(RowIndex, Index)
        If Not Entry Is Nothing Then
            Call Result.AddDataEntry(Entry)
        End If
    Next RowIndex
    Set Field = Result
End Function


Public Function Row(ByVal Index As Long) As RowEntry
    Dim Result      As RowEntry
    Dim FieldIndex  As Long
    Dim Entry       As DataEntry
    Dim FieldEntry  As FieldEntry

    If Index < 1 Or Index > mRowCount Then Exit Function
    Set Result = New RowEntry
    For FieldIndex = 1 To mFieldCount
        Set Entry = DataEntryAt(Index, FieldIndex)
        If Not Entry Is Nothing Then
            Set FieldEntry = New FieldEntry
            Set FieldEntry.Definition = mFieldDefinitions(FieldIndex)
            Set FieldEntry.DataEntry = Entry
            Call Result.AddFieldEntry(FieldEntry)
        End If
    Next FieldIndex
    Set Row = Result
End Function


Public Function Fields() As Collection
    Dim Result As Collection
    Dim FieldIndex As Long

    Set Result = New Collection

    For FieldIndex = 1 To mFieldCount
        Call Result.Add(Field(FieldIndex))
    Next FieldIndex

    Set Fields = Result
End Function


Public Function Rows() As Collection
    Dim Result As Collection
    Dim RowIndex As Long

    Set Result = New Collection

    For RowIndex = 1 To mRowCount
        Call Result.Add(Row(RowIndex))
    Next RowIndex

    Set Rows = Result
End Function


'============================================================
'====================[Adding]================================
'============================================================

Public Sub AddFieldDefinition(ByVal Value As FieldDefinition)
    Dim Definition As FieldDefinition
    Dim Temp       As New FieldDefinition

    Set Definition = Temp.Copy(Value)
    Call mFieldDefinitions.Add(Definition)
    mFieldCount = mFieldCount + 1
    Call ResizeDataEntries(mRowCount, mFieldCount)
End Sub


Public Sub AddEntry( ByVal Value As DataEntry, ByVal FieldIndex As Long, ByVal RowIndex As Long)
    Dim Entry As DataEntry
    If FieldIndex < 1 Or FieldIndex > mFieldCount Then
        Err.Raise vbObjectError + 1501,"DataSet.AddEntry","The specified FieldIndex does not exist."
    End If

    If RowIndex < 1 Or RowIndex > mRowCount Then
        Err.Raise vbObjectError + 1502,"DataSet.AddEntry","The specified RowIndex does not exist."
    End If

    Set Entry = mDataEntries(RowIndex, FieldIndex)

    If Entry Is Nothing Then
        Set Entry = New DataEntry
        Set mDataEntries(RowIndex, FieldIndex) = Entry
    End If

    Entry.Value = Value.Value
End Sub


Public Sub AddField(ByVal Value As Field)
    Dim Entry      As DataEntry
    Dim FieldIndex As Long
    Dim RowIndex   As Long

    Call AddFieldDefinition(Value.Definition)

    FieldIndex = mFieldCount

    For RowIndex = 1 To Value.Count
        Set Entry = Value.Item(RowIndex)
        If RowIndex > mRowCount Then
            Call AddEmptyRow()
        End If
        Call AddEntry(Entry, FieldIndex, RowIndex)
    Next RowIndex
End Sub


Public Sub AddRow(ByVal Value As RowEntry)
    Dim Entry      As FieldEntry
    Dim FieldIndex As Long
    Dim RowIndex   As Long

    Call AddEmptyRow()
    RowIndex = mRowCount
    For FieldIndex = 1 To Value.Count
        Set Entry = Value.Item(FieldIndex)
        If FieldIndex <= mFieldCount Then
            Call AddEntry(Entry.DataEntry, FieldIndex, RowIndex)
        End If
    Next FieldIndex
End Sub


Public Function Merge(ByVal DataSet1 As DataSet, ByVal DataSet2 As DataSet) As DataSet
    Dim Result     As DataSet
    Dim FieldIndex As Long
    Dim RowIndex   As Long
    Dim Entry      As DataEntry

    If DataSet1          Is Nothing           Then Err.Raise vbObjectError + 1504, "DataSet.Merge", "DataSet1 cannot be Nothing."
    If DataSet2          Is Nothing           Then Err.Raise vbObjectError + 1505, "DataSet.Merge", "DataSet2 cannot be Nothing."
    If DataSet1.RowCount <> DataSet2.RowCount Then Err.Raise vbObjectError + 1506, "DataSet.Merge", "The datasets must have the same number of rows."

    Set Result = New DataSet

    For FieldIndex = 1 To DataSet1.FieldCount
        Result.AddFieldDefinition DataSet1.FieldDefinition(FieldIndex)
    Next FieldIndex
    For FieldIndex = 1 To DataSet2.FieldCount
        Result.AddFieldDefinition DataSet2.FieldDefinition(FieldIndex)
    Next FieldIndex

    For RowIndex = 1 To DataSet1.RowCount
        Result.AddEmptyRow
    Next RowIndex

    '--------------------------------------------------------
    ' Copy DataSet1 entries.
    ' Create NEW DataEntry objects.
    '--------------------------------------------------------
    For RowIndex = 1 To DataSet1.RowCount
        For FieldIndex = 1 To DataSet1.FieldCount
            Set Entry = DataSet1.DataEntryAt(RowIndex, FieldIndex)
            If Not Entry Is Nothing Then
                Call Result.AddEntry(DataEntry.Copy(Entry), FieldIndex, RowIndex)
            End If
        Next FieldIndex
    Next RowIndex

    '--------------------------------------------------------
    ' Copy DataSet2 entries to the right.
    ' Create NEW DataEntry objects.
    '--------------------------------------------------------
    For RowIndex = 1 To DataSet2.RowCount
        For FieldIndex = 1 To DataSet2.FieldCount
            Set Entry = DataSet2.DataEntryAt(RowIndex, FieldIndex)

            If Not Entry Is Nothing Then
                Call Result.AddEntry(DataEntry.Copy(Entry), DataSet1.FieldCount + FieldIndex, RowIndex)
            End If
        Next FieldIndex
    Next RowIndex
    Set Merge = Result
End Function

Public Sub AddEmptyRow()
    mRowCount = mRowCount + 1
    Call ResizeDataEntries(mRowCount, mFieldCount)
End Sub

'============================================================
'====================[Helpers]===============================
'============================================================

Private Sub ResizeDataEntries(ByVal NewRowCount As Long, ByVal NewFieldCount As Long)
    Dim NewData()      As DataEntry
    Dim RowIndex       As Long
    Dim FieldIndex     As Long
    Dim CopyRows       As Long
    Dim CopyFields     As Long
    Dim OldRowCount    As Long
    Dim OldFieldCount  As Long

    If NewRowCount <= 0 Or NewFieldCount <= 0 Then
        Erase mDataEntries
        Exit Sub
    End If

    ReDim NewData(1 To NewRowCount, 1 To NewFieldCount)

    If GetDataEntryDimensions(mDataEntries, OldRowCount, OldFieldCount) Then
        CopyRows = MinLong(OldRowCount, NewRowCount)
        CopyFields = MinLong(OldFieldCount, NewFieldCount)

        For RowIndex = 1 To CopyRows
            For FieldIndex = 1 To CopyFields
                Set NewData(RowIndex, FieldIndex) = _
                    mDataEntries(RowIndex, FieldIndex)
            Next FieldIndex
        Next RowIndex
    End If

    mDataEntries = NewData
End Sub

Private Function GetDataEntryDimensions(ByRef Data() As DataEntry, ByRef RowCount As Long, ByRef FieldCount As Long) As Boolean
    On Error GoTo NotAllocated

    RowCount   = UBound(Data, 1) - LBound(Data, 1) + 1
    FieldCount = UBound(Data, 2) - LBound(Data, 2) + 1

    GetDataEntryDimensions = True
    Exit Function

    NotAllocated:
    RowCount = 0
    FieldCount = 0
    GetDataEntryDimensions = False
End Function



Private Function MinLong(ByVal Value1 As Long, ByVal Value2 As Long) As Long
    If Value1 < Value2 Then
        MinLong = Value1
    Else
        MinLong = Value2
    End If
End Function