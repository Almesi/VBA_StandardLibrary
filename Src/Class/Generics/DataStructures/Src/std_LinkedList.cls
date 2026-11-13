
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_LinkedList"
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

Public Function Create(Optional ByVal Value As Variant) As std_LinkedList
    Set Create = CreateTyped(Nothing, Value)
End Function

Public Function CreateTyped(ByVal Validator As IValidator, Optional ByVal Value As Variant) As std_LinkedList
    Set CreateTyped = New std_LinkedList
    Call CreateTyped.Initialize(Validator, Value)
End Function

Friend Sub Initialize(ByVal Validator As IValidator, ByVal Value As Variant)
    Set mValidator = Validator

    If Not IsMissing(Value) Then
        Set mFirstNode = std_Node.CreateTyped(mValidator, Value)
        Set mLastNode = mFirstNode
        mCount = 1
    End If
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

Public Property Get FirstNode() As std_Node
    Set FirstNode = mFirstNode
End Property

Public Property Get LastNode() As std_Node
    Set LastNode = mLastNode
End Property

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Set mEnumerator = std_NodeEnumerator.Create(mFirstNode, False)
    Set NewEnum = mEnumerator.NewEnum
End Property


' ==================================================
' ====================[Methods]=====================
' ==================================================

Public Sub Push(Optional ByVal Value As Variant)
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

Public Function Pop() As std_Node
    If IsEmpty Then Err.Raise vbObjectError + 1202, "std_LinkedList.Pop", "Linked list is empty."
    Set Pop = Remove(mCount)
End Function

Public Function Peek() As std_Node
    If IsEmpty Then Err.Raise vbObjectError + 1208, "std_LinkedList.Peek", "Linked list is empty."
    Set Peek = mLastNode
End Function

Public Sub Insert(ByVal Index As Long, Optional ByVal Value As Variant)
    Dim NewNode As std_Node
    Dim PreviousNode As std_Node
    Dim CurrentNode As std_Node

    If Index < 1 Or Index > mCount + 1 Then
        Err.Raise vbObjectError + 1204, "std_LinkedList.Insert", "Index is out of range."
    End If

    Set NewNode = std_Node.CreateTyped(mValidator, Value)

    If Index = 1 Then
        If Not mFirstNode Is Nothing Then
            Call NewNode.AddReference(mFirstNode)
        End If

        Set mFirstNode = NewNode

        If mCount = 0 Then Set mLastNode = NewNode
    Else
        Set PreviousNode = GetNode(Index - 1)

        If PreviousNode.HasReferences Then
            Set CurrentNode = PreviousNode.FirstReference
            Call NewNode.AddReference(CurrentNode)
            Call PreviousNode.ReplaceReference(CurrentNode, NewNode)
        Else
            Call PreviousNode.AddReference(NewNode)
        End If

        If Index = mCount + 1 Then Set mLastNode = NewNode
    End If

    mCount = mCount + 1
End Sub

Public Function Remove(ByVal Index As Long) As std_Node
    Dim PreviousNode As std_Node
    Dim CurrentNode As std_Node
    Dim NextNode As std_Node

    If Index < 1 Or Index > mCount Then
        Err.Raise vbObjectError + 1205, "std_LinkedList.Remove", "Index is out of range."
    End If

    Set CurrentNode = GetNode(Index)

    If CurrentNode.HasReferences Then
        Set NextNode = CurrentNode.FirstReference
    End If

    If Index = 1 Then
        Set mFirstNode = NextNode
    Else
        Set PreviousNode = GetNode(Index - 1)
        Call PreviousNode.RemoveReference(CurrentNode)

        If Not NextNode Is Nothing Then
            Call PreviousNode.AddReference(NextNode)
        End If
    End If

    Call CurrentNode.ClearReferences()

    If Index = mCount Then
        If Index = 1 Then
            Set mLastNode = Nothing
        Else
            Set mLastNode = PreviousNode
        End If
    End If

    mCount = mCount - 1

    Set Remove = CurrentNode
End Function

Public Function GetNode(Optional ByVal Index As Long = 0) As std_Node
    Dim CurrentNode As std_Node
    Dim CurrentIndex As Long

    If IsEmpty Then
        Err.Raise vbObjectError + 1206, "std_LinkedList.GetNode", "Linked list is empty."
    End If

    If Index = 0 Then Index = mCount

    If Index < 1 Or Index > mCount Then
        Err.Raise vbObjectError + 1207, "std_LinkedList.GetNode", "Index is out of range."
    End If

    If Index = 1 Then
        Set GetNode = mFirstNode
        Exit Function
    End If

    If Index = mCount Then
        Set GetNode = mLastNode
        Exit Function
    End If

    Set CurrentNode = mFirstNode

    For CurrentIndex = 1 To Index - 1
        Set CurrentNode = CurrentNode.FirstReference
    Next CurrentIndex

    Set GetNode = CurrentNode
End Function

Public Function IndexOf(ByVal Node As std_Node) As Long
    Dim CurrentNode As std_Node
    Dim CurrentIndex As Long

    If Node Is Nothing Then Exit Function

    Set CurrentNode = mFirstNode
    CurrentIndex = 1

    Do While Not CurrentNode Is Nothing
        If CurrentNode Is Node Then
            IndexOf = CurrentIndex
            Exit Function
        End If

        Set CurrentNode = CurrentNode.FirstReference
        CurrentIndex = CurrentIndex + 1
    Loop
End Function

Public Function ContainsNode(ByVal Node As std_Node) As Boolean
    ContainsNode = (IndexOf(Node) <> 0)
End Function

Public Sub Reverse()
    Dim CurrentNode As std_Node
    Dim PreviousNode As std_Node
    Dim NextNode As std_Node

    If mCount < 2 Then Exit Sub

    Set CurrentNode = mFirstNode
    Set mLastNode = mFirstNode

    Do While Not CurrentNode Is Nothing
        Set NextNode = CurrentNode.FirstReference
        Call CurrentNode.ClearReferences()

        If Not PreviousNode Is Nothing Then
            Call CurrentNode.AddReference(PreviousNode)
        End If

        Set PreviousNode = CurrentNode
        Set CurrentNode = NextNode
    Loop

    Set mFirstNode = PreviousNode
End Sub

Public Function ToArray() As Variant
    Dim Result() As Variant
    Dim CurrentNode As std_Node
    Dim i As Long

    If IsEmpty Then
        ToArray = Array()
        Exit Function
    End If

    ReDim Result(0 To mCount - 1)

    Set CurrentNode = mFirstNode

    For i = 0 To mCount - 1
        Call Assign(Result(i), CurrentNode.Value)
        Set CurrentNode = CurrentNode.FirstReference
    Next i

    ToArray = Result
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
