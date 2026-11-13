
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


' ==================================================
' ====================[Properties]==================
' ==================================================

Private mFirstNode  As std_Node
Private mLastNode   As std_Node
Private mCount      As Long
Private mValidator  As IValidator

Private mEnumerator As std_NodeEnumerator



' ==================================================
' ====================[Creation]====================
' ==================================================

Public Function Create(Optional ByVal Value As Variant) As std_Queue
    Set Create = CreateTyped(Nothing, Value)
End Function

Public Function CreateTyped(ByVal Validator As IValidator, Optional ByVal Value As Variant) As std_Queue
    Set CreateTyped = New std_Queue
    Call CreateTyped.Initialize(Validator, Value)
End Function

Friend Sub Initialize(ByVal Validator As IValidator, ByVal Value As Variant)
    Set mValidator = Validator
    Call Enqueue(Value)
End Sub


' ==================================================
' ====================[Getters/Setters]=============
' ==================================================

Public Property Get Count() As Long
    Count = mCount
End Property

Public Property Get IsEmpty() As Boolean
    IsEmpty = (mCount = 0)
End Property

Public Property Get FrontNode() As std_Node
    Set FrontNode = mFirstNode
End Property

Public Property Get BackNode() As std_Node
    Set BackNode = mLastNode
End Property

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Set mEnumerator = std_NodeEnumerator.Create(mFirstNode, False)
    Set NewEnum = mEnumerator.NewEnum
End Property


' ==================================================
' ====================[Methods]=====================
' ==================================================

Public Sub Enqueue(Optional ByVal Value As Variant)
    Dim NewNode As std_Node

    Set NewNode = std_Node.CreateTyped(mValidator, Value)

    If IsEmpty Then
        Set mFirstNode = NewNode
        Set mLastNode = NewNode
    Else
        Call mLastNode.AddReference(NewNode)
        Set mLastNode = NewNode
    End If

    mCount = mCount + 1
End Sub

Public Function Dequeue() As std_Node
    Dim CurrentNode As std_Node
    Dim NextNode As std_Node

    If IsEmpty Then
        Err.Raise vbObjectError + 1502, "std_Queue.Dequeue", "Queue is empty."
    End If

    Set CurrentNode = mFirstNode
    Set NextNode = CurrentNode.FirstReference

    Call CurrentNode.ClearReferences()

    Set mFirstNode = NextNode
    mCount = mCount - 1

    If IsEmpty Then
        Set mLastNode = Nothing
    End If

    Set Dequeue = CurrentNode
End Function

Public Function Peek() As std_Node
    If IsEmpty Then
        Err.Raise vbObjectError + 1503, "std_Queue.Peek", "Queue is empty."
    End If

    Set Peek = mFirstNode
End Function

Public Sub Clear()
    Dim CurrentNode As std_Node
    Dim NextNode As std_Node

    Set CurrentNode = mFirstNode

    Do While Not CurrentNode Is Nothing
        Set NextNode = CurrentNode.FirstReference
        Call CurrentNode.ClearReferences()
        Set CurrentNode = NextNode
    Loop

    Set mFirstNode = Nothing
    Set mLastNode = Nothing
    mCount = 0
End Sub


' ==================================================
' ====================[Helpers]=====================
' ==================================================
