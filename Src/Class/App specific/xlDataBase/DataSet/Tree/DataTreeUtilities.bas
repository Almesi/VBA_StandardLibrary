Attribute VB_Name = "DataTreeUtilities"


Option Explicit


Public Function TranslateTreeNodeToTreeDisplay(ByVal Source As DataTreeNode, ByVal NameField As Field) As TreeDisplay
    Dim Result       As TreeDisplay
    Dim Child        As DataTreeNode
    Dim ChildDisplay As TreeDisplay

    Dim NameEntry As DataEntry
    Set Result       = TreeDisplay.Create()
    Set Result.Value = Source.RowEntry

    Set NameEntry = NameField.Item(Source.RowEntry.Index)
    Result.DisplayName = CStr(NameEntry.Value)

    For Each Child In Source.Children
        Set ChildDisplay = TranslateTreeNodeToTreeDisplay(Child, NameField)
        Call Result.AddChild(ChildDisplay)
    Next Child

    Set TranslateTreeNodeToTreeDisplay = Result
End Function