VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "Query"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mCommand      As IQueryCommand
Private mSources      As Collection
Private mDataSet      As DataSet
Private mExecuted     As Boolean


'============================================================
'====================[Creation]==============================
'============================================================

Public Function Create(ByVal Command As IQueryCommand) As Query
    Set Create = New Query
    Call Create.Initialize(Command)
End Function


Friend Sub Initialize(ByVal Command As IQueryCommand)
    Set mCommand  = Command
    Set mSources  = New Collection
    Set mDataSet  = Nothing
    Let mExecuted = False
End Sub

'============================================================
'====================[Getters/Setters]=======================
'============================================================

Public Property Get Sources() As Collection
    Set Sources = mSources
End Property

Public Property Get SourceCount() As Long
    If mSources Is Nothing Then
        SourceCount = 0
    Else
        SourceCount = mSources.Count
    End If
End Property

Public Property Get Source(ByVal Index As Long) As DataSet
    Set Source = mSources(Index)
End Property

Public Property Set Command(ByVal Value As IQueryCommand)
    Set mCommand = Value
    Let mExecuted = False
End Property

Public Property Get Command() As IQueryCommand
    Set Command = mCommand
End Property

Public Property Get DataSet() As DataSet
    If Not mExecuted Then
        Err.Raise vbObjectError + 1104, _
                  "Query.DataSet", _
                  "Query has not been executed."
    End If
    Set DataSet = mDataSet
End Property

Public Property Get Executed() As Boolean
    Executed = mExecuted
End Property

'============================================================
'====================[Functions]=============================
'============================================================

Public Sub AddSource(ByVal NamedDataSet As NamedDataSet)
    If NamedDataSet.DataSet Is Nothing Then
        Err.Raise vbObjectError + 1100, _
                  "Query.AddSource", _
                  "Source DataSet cannot be Nothing."
    End If

    Call mSources.Add(NamedDataSet)
    Let mExecuted = False
End Sub

Public Function SourceByName(ByVal Name As String) As DataSet
    Dim Source As DataSet
    For Each Source In mSources
        If StrComp(Source.Name, Name, vbTextCompare) = 0 Then
            Set SourceByName = Source
            Exit Function
        End If
    Next Source
    Set SourceByName = Nothing
End Function


Public Sub Execute()
    If mCommand Is Nothing Then

        Err.Raise vbObjectError + 1101, _
                  "Query.Execute", _
                  "QueryCommand has not been assigned."
    End If
    If mSources Is Nothing Or mSources.Count = 0 Then
        Err.Raise vbObjectError + 1102, _
                  "Query.Execute", _
                  "Query has no source DataSets."
    End If

    Set mDataSet = mCommand.Execute(QueryContext.Create(mSources))

    If mDataSet Is Nothing Then
        Err.Raise vbObjectError + 1103, _
                  "Query.Execute", _
                  "QueryCommand returned no DataSet."
    End If
    Let mExecuted = True
End Sub

Public Sub Clear()
    Set mSources  = New Collection
    Set mDataSet  = Nothing
    Let mExecuted = False
End Sub

'============================================================
'====================[Helpers]===============================
'============================================================