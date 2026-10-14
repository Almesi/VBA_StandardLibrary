VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mNames  As New Collection
Private mValues As New Collection

Public Function Create() As Settings
    Set Create = New Settings
End Function

Public Sub Add(ByVal nName As String, ByVal nValue As IGetValue)
    If Exists(nName) Then
        Call SetValue(nName, nValue)
        Exit Sub
    End If

    Call mNames.Add(nName)
    Call mValues.Add(nValue)
End Sub

Public Sub SetValue(ByVal nName As String, ByVal nValue As IGetValue)
    Dim Found As Long
    Found = Find(nName)
    If Found <> 0 Then
        Set mValues(Found) = nValue
      Else
          Call Add(nName, nValue)
    End If
End Sub

Public Function GetValue(ByVal nName As String) As IGetValue
    Dim Found As Long
    Found = Find(nName)
    If Found <> 0 Then
        Set GetValue = mValues(Found)
        Exit Function
    End If
    Err.Raise 5, TypeName(Me), "Setting not found: '" & nName & "'"
End Function

Public Function TryGetValue(ByVal nName As String, ByRef nValue As IGetValue) As Boolean
    Dim Found As Long
    Found = Find(nName)
    If Found <> 0 Then
        Set nValue = mValues(Found)
        TryGetValue = True
    End If
End Function

Public Sub Remove(ByVal nName As String)
    Dim Found As Long
    Found = Find(nName)
    If Found <> 0 Then
        Call mNames.Remove(Found)
        Call mValues.Remove(Found)
    End If
End Sub

Public Sub Clear()
    Do While mNames.Count > 0
        Call mNames.Remove(mNames.Count)
        Call mValues.Remove(mValues.Count)
    Loop
End Sub

Public Function Exists(ByVal nName As String) As Boolean
    Exists = Find(nName) <> 0
End Function

Public Function Find(ByVal nName As String) As Long
    Dim i As Long

    For i = 1 To mNames.Count
        If StrComp(CStr(mNames(i)), nName, vbTextCompare) = 0 Then
            Find = i
            Exit Function
        End If
    Next
    Find = 0
End Function

Public Property Get Count() As Long
    Count = mNames.Count
End Property

Public Function Name(ByVal Index As Long) As String
    Name = CStr(mNames(Index))
End Function

Public Function Value(ByVal Index As Long) As IGetValue
    Set Value = mValues(Index)
End Function

Public Property Get Names() As Collection
    Set Names = mNames
End Property

Public Property Get Values() As Collection
    Set Values = mValues
End Property