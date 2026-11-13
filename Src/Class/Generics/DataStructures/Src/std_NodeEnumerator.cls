
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_NodeEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


' ==================================================
' ====================[Properties]==================
' ==================================================

Private mNode       As std_Node
Private mCollection As New Collection


' ==================================================
' ====================[Creation]====================
' ==================================================

Public Function Create(ByVal nNode As std_Node, Optional ByVal DepthFirst As Boolean = True) As std_NodeEnumerator
    Set Create = New std_NodeEnumerator
    Call Create.Initialize(nNode, DepthFirst)
End Function

Friend Sub Initialize(ByVal nNode As std_Node, Optional ByVal DepthFirst As Boolean = True)
    Set mNode = nNode

    If DepthFirst Then
        Call BuildDepthFirst(mNode)
    Else
        Call BuildBreadthFirst(mNode)
    End If
End Sub


' ==================================================
' ====================[Getters/Setters]=============
' ==================================================

Public Property Get NewEnum() As IUnknown
    Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = mCollection.[_NewEnum]
End Property


' ==================================================
' ====================[Helpers]=====================
' ==================================================

Private Sub BuildDepthFirst(ByVal nNode As std_Node)
    Dim i           As Long
    Dim CurrentNode As std_Node

    If nNode Is Nothing Then Exit Sub
    If ContainsNode(nNode) Then Exit Sub

    Call mCollection.Add(nNode)

    For i = 1 To nNode.ReferenceCount
        Set CurrentNode = nNode.GetReference(i)
        Call BuildDepthFirst(CurrentNode)
    Next i
End Sub

Private Sub BuildBreadthFirst(ByVal nNode As std_Node)
    Dim PendingNodes  As New Collection
    Dim CurrentNode   As std_Node
    Dim ReferenceNode As std_Node
    Dim i As Long

    If nNode Is Nothing Then Exit Sub

    Call PendingNodes.Add(nNode)

    Do While PendingNodes.Count > 0
        Set CurrentNode = PendingNodes(1)
        Call PendingNodes.Remove(1)

        If Not ContainsNode(CurrentNode) Then
            Call mCollection.Add(CurrentNode)

            For i = 1 To CurrentNode.ReferenceCount
                Set ReferenceNode = CurrentNode.GetReference(i)

                If Not ContainsNode(ReferenceNode) Then
                    Call PendingNodes.Add(ReferenceNode)
                End If
            Next i
        End If
    Loop
End Sub

Private Function ContainsNode(ByVal nNode As std_Node) As Boolean
    Dim CurrentNode As std_Node

    For Each CurrentNode In mCollection
        If CurrentNode Is nNode Then
            ContainsNode = True
            Exit Function
        End If
    Next CurrentNode
End Function