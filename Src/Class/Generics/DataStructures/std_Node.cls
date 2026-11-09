VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ==================================================
' ====================[Properties]==================
' ==================================================

Private mValue      As Variant
Private mReferences As std_Collection


' ==================================================
' ====================[Creation]====================
' ==================================================

Public Function Create(Optional ByVal Value As Variant) As std_Node
    Set Create = New std_Node
    Call Create.Initialize(Value)
End Function

Public Function Clone() As std_Node
    Set Clone = Create(mValue)
    Call Clone.References.AddFromstd_Collection(mReferences)
End Function

Public Function CloneDeep() As std_Node
    Dim Item As std_Node

    Set CloneDeep = Create(mValue)

    For Each Item In mReferences
        Call CloneDeep.AddReference(Item.CloneDeep)
    Next Item
End Function

Friend Sub Initialize(Optional ByVal Value As Variant)
    Call AssignValue(mValue, Value)
End Sub

Private Sub Class_Initialize()
    Set mReferences = std_Collection.CreateFromVariant(std_Node)
End Sub


' ==================================================
' ====================[Properties]==================
' ==================================================

Public Property Get ID() As LongPtr
    ID = ObjPtr(Me)
End Property

Public Property Get Value() As Variant
    Call Assign(Value, mValue)
End Property

Public Property Let Value(ByVal NewValue As Variant)
    mValue = NewValue
End Property

Public Property Set Value(ByVal NewValue As Object)
    Set mValue = NewValue
End Property

Public Sub AssignValue(ByVal NewValue As Variant)
    Call Assign(mValue, NewValue)
End Sub

Public Property Get References() As std_Collection
    Set References = mReferences
End Property

Public Property Get ReferenceCount() As Long
    ReferenceCount = mReferences.Count
End Property

Public Property Get HasReferences() As Boolean
    HasReferences = Not mReferences.IsEmpty
End Property

Public Property Get HasValue() As Boolean
    HasValue = Not IsMissing(mValue)
End Property

Public Property Get FirstReference() As std_Node
    If HasReferences Then
        Set FirstReference = mReferences.First
    End If
End Property

Public Property Get LastReference() As std_Node
    If HasReferences Then
        Set LastReference = mReferences.Last
    End If
End Property

Public Property Get IsReferencedBy(ByVal Node As std_Node) As Boolean
    If Node Is Nothing Then Exit Property
    IsReferencedBy = Node.HasReference(Me)
End Property


' ==================================================
' ====================[Methods]=====================
' ==================================================

Public Sub AddReference(ByVal Node As std_Node)
    Call ValidateNode(Node)

    If Not HasReference(Node) Then
        Call mReferences.Add(Node)
    End If
End Sub

Public Sub InsertReference(ByVal Node As std_Node, ByVal Index As Long)
    Call ValidateNode(Node)

    If HasReference(Node) Then Exit Sub
    Call mReferences.Insert(Node, Index)
End Sub

Public Sub RemoveReference(ByVal Node As std_Node)
    Dim Index As Long
    Call ValidateNode(Node)

    Index = mReferences.IndexOf(Node)
    If Index = 0 Then Exit Sub
    Call mReferences.Remove(Index)
End Sub

Public Sub RemoveReferenceAt(ByVal Index As Long)
    If Not mReferences.Exists(Index) Then
        Err.Raise vbObjectError + 1100, "std_Node.RemoveReferenceAt", "Reference index is out of range."
    End If

    Call mReferences.Remove(Index)
End Sub

Public Function GetReference(ByVal Index As Long) As std_Node
    If Not mReferences.Exists(Index) Then
        Err.Raise vbObjectError + 1101, "std_Node.GetReference", "Reference index is out of range."
    End If

    Set GetReference = mReferences.Item(Index)
End Function

Public Function GetReferenceIndex(ByVal Node As std_Node) As Long
    Call ValidateNode(Node)
    GetReferenceIndex = mReferences.IndexOf(Node)
End Function

Public Function HasReference(ByVal Node As std_Node, Optional ByVal Recursive As Boolean = False) As Boolean
    Dim Visited As New Collection

    If Node Is Nothing Then Exit Function

    HasReference = mReferences.Contains(Node)

    If Not Recursive Or HasReference Then Exit Function

    HasReference = HasReferenceRecursive(Node, Visited)
End Function

Public Function HasReferenceAt(ByVal Index As Long) As Boolean
    HasReferenceAt = mReferences.Exists(Index)
End Function

Public Function ContainsReference(ByVal Node As std_Node) As Boolean
    ContainsReference = HasReference(Node)
End Function

Public Sub ReplaceReference(ByVal OldNode As std_Node, ByVal NewNode As std_Node)
    Dim Index As Long

    Call ValidateNode(OldNode)
    Call ValidateNode(NewNode)

    Index = mReferences.IndexOf(OldNode)

    If Index = 0             Then Err.Raise vbObjectError + 1102, "std_Node.ReplaceReference", "Reference does not exist."
    If HasReference(NewNode) Then Err.Raise vbObjectError + 1103, "std_Node.ReplaceReference", "New reference already exists."

    Call mReferences.Remove(Index)
    Call mReferences.Insert(NewNode, Index)
End Sub

Public Sub MoveReference(ByVal OldIndex As Long, ByVal NewIndex As Long)
    Dim Node As std_Node

    If Not mReferences.Exists(OldIndex)             Then Err.Raise vbObjectError + 1104, "std_Node.MoveReference", "Source index is out of range."
    If NewIndex < 1 Or NewIndex > mReferences.Count Then Err.Raise vbObjectError + 1105, "std_Node.MoveReference", "Destination index is out of range."

    If OldIndex = NewIndex Then Exit Sub

    Set Node = mReferences.Item(OldIndex)

    Call mReferences.Remove(OldIndex)
    Call mReferences.Insert(Node, NewIndex)

End Sub

Public Sub ClearReferences()
    Call mReferences.Clear
End Sub

Public Function GetByID(ByVal ID As LongPtr, Optional ByVal Recursive As Boolean = False) As std_Node
    Dim i As Long
    Dim Node As std_Node

    For i = 1 To ReferenceCount
        Set Node = GetReference(i)

        If Node.ID = ID Then
            Set GetByID = Node
            Exit Function
        End If

        If Recursive Then
            Set GetByID = Node.GetByID(ID, True)

            If Not GetByID Is Nothing Then Exit Function
        End If
    Next i
End Function

Public Function GetParent(ByVal Node As std_Node, Optional ByVal Recursive As Boolean = True) As std_Node
    Dim CurrentNode As std_Node
    Dim i As Long

    If Node Is Nothing Then Exit Function

    For i = 1 To ReferenceCount
        Set CurrentNode = GetReference(i)

        If CurrentNode Is Node Then
            Set GetParent = Me
            Exit Function
        End If

        If Recursive Then
            Set GetParent = CurrentNode.GetParent(Node, True)

            If Not GetParent Is Nothing Then Exit Function
        End If
    Next i
End Function

Public Function GetDescendantCount() As Long
    Dim Node As std_Node

    For Each Node In mReferences
        GetDescendantCount = GetDescendantCount + 1 + Node.GetDescendantCount
    Next Node
End Function

' ==================================================
' ====================[Helpers]=====================
' ==================================================

Private Sub ValidateNode(ByVal Node As std_Node)
    If Node Is Nothing Then Err.Raise vbObjectError + 1106, "std_Node", "Node cannot be Nothing."
End Sub

Private Function HasReferenceRecursive(ByVal Node As std_Node, ByVal Visited As Collection) As Boolean
    Dim CurrentNode As std_Node
    Dim i As Long

    If Node Is Nothing Then Exit Function
    If IsVisited(Node, Visited) Then Exit Function

    Call Visited.Add(Node)

    For i = 1 To ReferenceCount
        Set CurrentNode = GetReference(i)

        If CurrentNode Is Node Then
            HasReferenceRecursive = True
            Exit Function
        End If

        If CurrentNode.HasReferenceRecursive(Node, Visited) Then
            HasReferenceRecursive = True
            Exit Function
        End If
    Next i
End Function

Private Function IsVisited(ByVal Node As std_Node, ByVal Visited As Collection) As Boolean
    Dim Item As std_Node

    For Each Item In Visited
        If Item Is Node Then
            IsVisited = True
            Exit Function
        End If
    Next Item
End Function