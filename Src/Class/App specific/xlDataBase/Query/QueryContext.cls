VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mSources As Collection
Private mCurrent As DataSet


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal Sources As Collection) As QueryContext
    Set Create = New QueryContext
    Call Create.Initialize(Sources)
End Function

Friend Sub Initialize(ByVal Sources As Collection)
    Set mSources = New Collection
    Set mCurrent = Nothing
    Call CopySources(Sources)
End Sub


Private Sub CopySources(ByVal Sources As Collection)
    Dim Source As NamedDataSet
    If Sources Is Nothing Then Exit Sub

    For Each Source In Sources
        Call mSources.Add(Source)
    Next Source
End Sub


Public Sub Add(ByVal Value As NamedDataSet)
    Call mSources.Add(Value)
End Sub

'===============================================================================
' SOURCES
'===============================================================================

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


Public Property Get Source(ByVal Index As Long) As NamedDataSet
    Set Source = mSources(Index)
End Property


Public Function SourceByName(ByVal Name As String) As NamedDataSet
    Dim Source As NamedDataSet
    For Each Source In mSources
        If StrComp(Source.Name, Name, vbTextCompare) = 0 Then
            Set SourceByName = Source
            Exit Function
        End If
    Next Source
    Set SourceByName = Nothing
End Function


Public Function DataSetByName(ByVal Name As String) As DataSet
    Dim Source As NamedDataSet
    Set Source = SourceByName(Name)

    If Source Is Nothing Then
        Set DataSetByName = Nothing
    Else
        Set DataSetByName = Source.DataSet
    End If
End Function


'===============================================================================
' CURRENT DATASET
'===============================================================================

Public Property Get Current() As DataSet
    Set Current = mCurrent
End Property


Public Property Set Current(ByVal Value As DataSet)
    Set mCurrent = Value
End Property


Public Function HasCurrent() As Boolean
    HasCurrent = Not mCurrent Is Nothing
End Function


'===============================================================================
' RESET
'===============================================================================

Public Sub Reset()
    Set mCurrent = Nothing
End Sub