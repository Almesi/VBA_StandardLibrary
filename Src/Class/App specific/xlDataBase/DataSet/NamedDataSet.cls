VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "NamedDataSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mName    As String
Private mDataSet As DataSet

'============================================================
'====================[Creation]==============================
'============================================================

Public Function Create(ByVal Name As String, ByVal DataSet As DataSet) As NamedDataSet
    Set Create = New NamedDataSet
    Call Create.Initialize(Name, DataSet)
End Function

Friend Sub Initialize(ByVal Name As String, ByVal DataSet As DataSet)
    If Len(Trim(Name)) = 0 Then
        Err.Raise vbObjectError + 1200, _
                  "NamedDataSet.Initialize", _
                  "Source name cannot be empty."
    End If

    If DataSet Is Nothing Then
        Err.Raise vbObjectError + 1201, _
                  "NamedDataSet.Initialize", _
                  "Source DataSet cannot be Nothing."
    End If

    Let mName = Name
    Set mDataSet = DataSet
End Sub


'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Get Name() As String
    Name = mName
End Property

Public Property Get DataSet() As DataSet
    Set DataSet = mDataSet
End Property

'============================================================
'====================[Functions]=============================
'============================================================