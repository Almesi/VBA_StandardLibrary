VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'============================================================
'====================[Properties]============================
'============================================================


Private mTables      As Object
Private mDataSets    As Object
Private mClassifier  As FieldClassifier
Private mInitialized As Boolean

'============================================================
'====================[Creation]==============================
'============================================================

Public Function Create(Optional ByVal Classifier As FieldClassifier = Nothing) As DataBase
    Set Create = New DataBase
    Call Create.Initialize(Classifier)
End Function

Friend Sub Initialize(Optional ByVal Classifier As FieldClassifier = Nothing)
    Set mTables = CreateObject("Scripting.Dictionary")
    mTables.CompareMode = vbTextCompare

    Set mDataSets = CreateObject("Scripting.Dictionary")
    mDataSets.CompareMode = vbTextCompare

    Set mClassifier = Classifier
    mInitialized = False

    Call LoadTables()

    mInitialized = True
End Sub


'============================================================
'====================[Functions]=============================
'============================================================

Private Sub LoadTables()
    Dim WS As Worksheet
    Dim Table As Table
    Dim DataSet As DataSet

    For Each WS In ThisWorkbook.Worksheets
        If IsDatabaseWorksheet(WS) Then
            Set Table = CreateTableFromSheet(WS)
            If Not Table Is Nothing Then
                Set DataSet = Table.DataSet
                Call mTables.Add(WS.Name, Table)
                Call mDataSets.Add(WS.Name, DataSet)
            End If
        End If
    Next WS
End Sub

Public Sub Remove(ByVal Name As String)
    If mTables Is Nothing Then Exit Sub
    If mTables.Exists(Name) Then
        Call mTables.Remove(Name)
    End If
    If Not mDataSets Is Nothing Then
        If mDataSets.Exists(Name) Then
            Call mDataSets.Remove(Name)
        End If
    End If
End Sub

Public Sub Clear()
    If Not mTables Is Nothing Then
        Call mTables.RemoveAll()
    End If
    If Not mDataSets Is Nothing Then
        Call mDataSets.RemoveAll()
    End If
    mInitialized = False
End Sub


Public Function GetTable(ByVal Name As String) As Table
    If Not mInitialized Then
        Err.Raise vbObjectError + 1000, _
                  "Database.GetTable", _
                  "Database has not been initialized."

    End If

    If Len(Trim(Name)) = 0 Then
        Err.Raise vbObjectError + 1001, _
                  "Database.GetTable", _
                  "Table name cannot be empty."

    End If

    If Not mTables.Exists(Name) Then
        Err.Raise vbObjectError + 1002, _
                  "Database.GetTable", _
                  "Database table '" & Name & "' was not found."

    End If

    Set GetTable = mTables(Name)
End Function


Public Function GetDataSet(ByVal Name As String) As DataSet
    If Not mInitialized Then
        Err.Raise vbObjectError + 1010, _
                  "Database.GetDataSet", _
                  "Database has not been initialized."
    End If

    If Len(Trim(Name)) = 0 Then
        Err.Raise vbObjectError + 1011, _
                  "Database.GetDataSet", _
                  "DataSet name cannot be empty."
    End If

    If Not mDataSets.Exists(Name) Then
        Err.Raise vbObjectError + 1012, _
                  "Database.GetDataSet", _
                  "Database DataSet '" & Name & "' was not found."
    End If

    Set GetDataSet = mDataSets(Name)
End Function


Public Function GetWrite(ByVal Name As String) As TableWrite
    Dim Table As Table
    Dim Writer As TableWrite

    Set Table = GetTable(Name)
    Set Writer = New TableWrite
    Set Writer.Table = Table
    Set GetWrite = Writer
End Function


Public Sub Refresh(Optional ByVal Name As String = Empty)
    Dim Table As Table

    If Len(Trim(Name)) = 0 Then
        Call RefreshAll()
    Else
        Set Table = GetTable(Name)
        Call Table.Refresh()
        Set mDataSets(Name) = Table.DataSet
    End If
End Sub


Private Sub RefreshAll()
    Dim Key As Variant
    Dim Table As Table

    For Each Key In mTables.Keys
        Call Table.Refresh(Key)
    Next Key
End Sub


Public Function Exists(ByVal Name As String) As Boolean
    If mTables Is Nothing Then
        Exists = False
        Exit Function
    End If
    Exists = mTables.Exists(Name)
End Function

Public Property Get Count() As Long
    If mTables Is Nothing Then
        Count = 0
    Else
        Count = mTables.Count
    End If
End Property

Public Function Names() As Variant
    If mTables Is Nothing Then
        Names = Array()
    Else
        Names = mTables.Keys
    End If
End Function


'============================================================
'====================[Getters/Setters]=======================
'============================================================

Public Property Get Initialized() As Boolean
    Initialized = mInitialized
End Property

Public Property Set Classifier(ByVal Value As FieldClassifier)
    Set mClassifier = Value
End Property

Public Property Get Classifier() As FieldClassifier
    Set Classifier = mClassifier
End Property




'============================================================
'====================[Helpers]===============================
'============================================================

Private Function CreateTableFromSheet(ByVal WS As Worksheet) As Table
    Dim Result As Table
    Set Result = Table.CreateFromTable(WS.ListObjects(WS.Name), mClassifier)
    Set CreateTableFromSheet = Result
End Function

Private Function IsDatabaseWorksheet(ByVal WS As Worksheet) As Boolean
    IsDatabaseWorksheet = (UCase(Left(WS.Name, 3)) = "DB_")
End Function