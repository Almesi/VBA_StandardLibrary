
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


' ==================================================
' ====================[Properties]==================
' ==================================================

Private mTopNode    As std_Node
Private mCount      As Long

Private mEnumerator As std_NodeEnumerator


' ==================================================
' ====================[Creation]====================
' ==================================================

Public Function Create() As std_Stack
    Set Create = New std_Stack
End Function

Public Function CreateWithValue(Optional ByVal Value As Variant) As std_Stack
    Set CreateWithValue = Create()
    Call CreateWithValue.Push(Value)
End Function


' ==================================================
' ====================[Getters/Setters]=============
' ==================================================

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = (mCount = 0)
End Property

Public Property Get TopNode() As std_Node
    Set TopNode = mTopNode
End Property

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Set mEnumerator = std_NodeEnumerator.Create(mTopNode, False)
    Set NewEnum = mEnumerator.[_NewEnum]
End Property


' ==================================================
' ====================[Methods]=====================
' ==================================================

Public Sub Push(Optional ByVal Value As Variant)
    Dim NewNode As std_Node

    Set NewNode = std_Node.Create(Value)

    If Not mTopNode Is Nothing Then
        Call NewNode.AddReference(mTopNode)
    End If

    Set mTopNode = NewNode
    mCount = mCount + 1
End Sub

Public Function Pop() As std_Node
    Dim CurrentNode As std_Node
    Dim NextNode As std_Node

    If IsEmpty Then
        Err.Raise vbObjectError + 1401, "std_Stack.Pop", "Stack is empty."
    End If

    Set CurrentNode = mTopNode
    Set NextNode = CurrentNode.FirstReference

    Call CurrentNode.ClearReferences()

    Set mTopNode = NextNode
    mCount = mCount - 1

    Set Pop = CurrentNode
End Function

Public Function Peek() As std_Node
    If IsEmpty Then
        Err.Raise vbObjectError + 1402, "std_Stack.Peek", "Stack is empty."
    End If

    Set Peek = mTopNode
End Function

Public Sub Clear()
    Dim CurrentNode As std_Node
    Dim NextNode As std_Node

    Set CurrentNode = mTopNode

    Do While Not CurrentNode Is Nothing
        Set NextNode = CurrentNode.FirstReference
        Call CurrentNode.ClearReferences()
        Set CurrentNode = NextNode
    Loop

    Set mTopNode = Nothing
    mCount = 0
End Sub


' ==================================================
' ====================[Helpers]=====================
' ==================================================
