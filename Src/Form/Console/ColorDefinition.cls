VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Private p_Colors() As Long
Private p_Names() As String
Private Count As Long

Private Sub Class_Initialize()
    Count = -1
End Sub

Public Sub AddColor(Name As String, Color As Long)
    Count = Count + 1
    ReDim Preserve p_Colors(Count)
    ReDim Preserve p_Names(Count)
    p_Colors(Count) = Color
    p_Names(Count) = Name
End Sub

Public Property Get Colors() As Long()
    Colors = p_Colors
End Property
Public Property Get Names() As String()
    Names = p_Names
End Property

Public Property Get Name(Index As Long) As String
    Name = p_Names(Index)
End Property
Public Property Let Name(Index As Long, n_Value As String)
    p_Names(Index) = Name
End Property

Public Property Get Color(Name As String) As Long
    Dim i As Long
    Color = -1
    For i = 0 To Count
        If p_Names(i) = Name Then
            Color = ColorByIndex(i)
            Exit For
        End If
    Next i
End Property
Public Property Let Color(Name As String, Value As Long)
    Dim i As Long
    For i = 0 To Count
        If p_Names(i) = Name Then
            ColorByIndex(i) = Value
            Exit For
        End If
    Next i
End Property

Public Property Get ColorByIndex(Index As Long) As Long
    ColorByIndex = p_Colors(Index)
End Property
Public Property Let ColorByIndex(Index As Long, Value As Long)
    p_Colors(Index) = Value
End Property