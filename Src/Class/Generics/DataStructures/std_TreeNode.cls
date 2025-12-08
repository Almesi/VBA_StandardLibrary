    VERSION 1.0 CLASS
    BEGIN
    MultiUse = -1  'True
    END
    Attribute VB_Name = "std_TreeNode"
    Attribute VB_GlobalNameSpace = False
    Attribute VB_Creatable = False
    Attribute VB_PredeclaredId = True
    Attribute VB_Exposed = True


Option Explicit

Private p_Value As Variant
Private p_Branches() As Long
Private p_Size As Long

Public Property Let Value(n_Value As Variant)
    If IsObject(n_Value) Then
        Set p_Value = n_Value
    Else
        p_Value = n_Value
    End If
End Property
Public Property Get Value() As Variant
    If IsObject(p_Value) Then
        Set Value = p_Value
    Else
        Value = p_Value
    End If
End Property

Public Property Let Branches(n_Value() As Long)
    p_Branches = n_Value
    p_Size = Ubound(n_Value)
End Property
Public Property Get Branches() As Long()
    Branches = p_Branches
End Property

Public Property Let Branch(Index As Long, n_Value As Long)
    p_Branches(Index) = n_Value
End Property
Public Property Get Branch(Index As Long) As Long
    Branch = p_Branches(Index)
End Property

Public Function AddBranch(Value As Long)
    p_Size = p_Size + 1
    ReDim Preserve p_Branches(p_Size)
    p_Branches(p_Size) = Value
End Function

Private Sub Class_Initialize()
    p_Size = -1
End Sub