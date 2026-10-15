VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private mRoots As New Collection

Public Function Create(ByVal Source As DataSet, ByVal ParentField As Field) As DataTree
    Set Create = New DataTree
    Call Create.Initialize(Source, ParentField)
End Function

Public Property Get Roots() As Collection
    Set Roots = mRoots
End Property

Public Property Get RootCount() As Long
    RootCount = mRoots.Count
End Property

Friend Sub Initialize(ByVal Source As DataSet, ByVal ParentField As Field)
    Dim Nodes()    As DataTreeNode
    Dim Node       As DataTreeNode
    Dim ParentNode As DataTreeNode

    Dim ParentEntry As DataEntry
    Dim RowIndex    As Long
    Dim ParentIndex As Long

    If Source      Is Nothing Then Err.Raise vbObjectError + 1601, "DataSetTree.Build", "Source cannot be Nothing."
    If ParentField Is Nothing Then Err.Raise vbObjectError + 1602, "DataSetTree.Build", "ParentField cannot be Nothing."


    ReDim Nodes(1 To Source.RowCount)

    ' First pass: create one node per row.
    For RowIndex = 1 To Source.RowCount
        Set Node = DataTreeNode.Create(Source.Row(RowIndex))
        Set Nodes(RowIndex) = Node
    Next RowIndex

    ' Second pass: connect nodes to their parents.
    For RowIndex = 1 To Source.RowCount
        Set ParentEntry = ParentField.Item(RowIndex)

        ' Missing parent entry means this row is a root.
        If ParentEntry Is Nothing Then
            Call mRoots.Add(Nodes(RowIndex))
        Else
            ParentIndex = CLng(ParentEntry.Value)
                If ParentIndex < 0 Or ParentIndex > Source.RowCount Then
                    Err.Raise vbObjectError + 1604, "DataSetTree.Build", "Parent row index out of range at row " & RowIndex
                ElseIf ParentIndex = RowIndex Then
                    Err.Raise vbObjectError + 1605, "DataSetTree.Build", "A row cannot be its own parent."
                Else
                    Set ParentNode = Nodes(ParentIndex)
                    Call ParentNode.AddChild(Nodes(RowIndex))
                End If
            End If
        End If
    Next RowIndex

End Sub
