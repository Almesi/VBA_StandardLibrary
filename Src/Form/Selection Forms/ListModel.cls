VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "ListModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    Items    As Collection
    Selected As Collection
End Type

Private This As TModel

Private Sub Class_Initialize()
    Set This.Items = New Collection
    Set This.Selected = New Collection
End Sub


Public Property Get Items() As Collection
    Set Items = This.Items
End Property


Public Sub Add(ByVal Value As Variant)
    Call This.Items.Add(Value)
End Sub


Public Sub Clear()
    Set This.Items = New Collection
    Set This.Selected = New Collection
End Sub


Public Property Get Count() As Long
    Count = This.Items.Count
End Property


Public Property Get Selected() As Collection
    Set Selected = This.Selected
End Property


Public Sub SelectItem(ByVal Index As Long)
    If Index < 1 Then Exit Sub
    If Index > This.Items.Count Then Exit Sub
    Call This.Selected.Add(Index)
End Sub


Public Property Get IsSelected(ByVal Index As Long) As Boolean
    Dim Item As Variant
    For Each Item In This.Selected
        If Item = Index Then
            IsSelected = True
            Exit Property
        End If
    Next Item
End Property


Public Property Get Self() As ListModel
    Set Self = Me
End Property