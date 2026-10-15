VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TreeValueForm 
   Caption         =   "Select Value from Tree"
   ClientHeight    =   4665
   ClientWidth     =   8460.001
   OleObjectBlob   =   "TreeValueForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TreeValueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IView

Private Type TView
    ReturnValue As FormReturn
    Model       As ValueModel
End Type

Private This            As TView
Private DisplayTree     As TreeDisplay
Private DisplayNodes    As Collection
Private SelectedNodeKey As String

Private TreeDepth       As Long
Private LongestName     As Long

Public Function Create(ByVal nTree As TreeDisplay) As TreeValueForm
    Set Create = New TreeValueForm
    Call Create.Initialize(nTree)
End Function

Friend Sub Initialize(ByVal nTree As TreeDisplay)
    Set DisplayTree = nTree
End Sub

Private Sub PopulateTree()
    Call ValueTree.Nodes.Clear
    Set DisplayNodes = New Collection
    SelectedNodeKey = vbNullString
    TreeDepth = 0
    LongestName = 0

    If DisplayTree Is Nothing Then
        Call ResizeForm
        Exit Sub
    End If

    Call AddDisplayNode(DisplayTree, vbNullString, "1", 1)
    If ValueTree.Nodes.Count > 0 Then ValueTree.Nodes(1).Expanded = True
    Call ResizeForm
End Sub

Private Sub AddDisplayNode(ByVal Display As TreeDisplay, ByVal ParentKey As String, ByVal Path As String, ByVal Depth As Long)
    Dim NodeKey As String
    Dim Child   As TreeDisplay
    Dim i       As Long

    NodeKey = "N" & Path

    If Len(ParentKey) = 0 Then
        Call ValueTree.Nodes.Add(, , NodeKey, Display.DisplayName)
    Else
        Call ValueTree.Nodes.Add(ParentKey, MSComctlLib.tvwChild, NodeKey, Display.DisplayName)
    End If

    Call DisplayNodes.Add(Display, NodeKey)

    If Depth > TreeDepth Then TreeDepth = Depth
    If Len(Display.DisplayName) > LongestName Then LongestName = Len(Display.DisplayName)

    For i = 1 To Display.Children.Count
        Set Child = Display.Children(i)
        Call AddDisplayNode(Child, NodeKey, Path & "." & CStr(i), Depth + 1)
    Next i
End Sub

Private Sub OKButton_Click()
    Dim SelectedDisplay As TreeDisplay

    If Len(SelectedNodeKey) = 0 Then Exit Sub
    Set SelectedDisplay = DisplayNodes.Item(SelectedNodeKey)
    Call This.Model.Assign(SelectedDisplay)
    This.ReturnValue = FormReturnOK
    Me.Hide
End Sub

Private Sub CancelButton_Click()
    Call OnCancel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Call OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.ReturnValue = FormReturnCancel
    Me.Hide
End Sub

Private Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model   = Model
    This.ReturnValue = FormReturnCancel
    Call PopulateTree
    Me.Show vbModal
    IView_Show = This.ReturnValue
End Function

Private Property Let IView_Result(ByVal Value As FormReturn)
    This.ReturnValue = Value
End Property

Private Property Get IView_Form() As Object
    Set IView_Form = Me
End Property

Private Sub ValueTree_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node Is Nothing Then Exit Sub
    SelectedNodeKey = Node.Key
    ResizeForm
End Sub

Private Sub ValueTree_Expand(ByVal Node As MSComctlLib.Node)
    ResizeForm
End Sub

Private Sub ValueTree_Collapse(ByVal Node As MSComctlLib.Node)
    ResizeForm
End Sub



' Resizing

Private Sub ResizeForm()
    Const ROW_HEIGHT     As Long = 20
    Const INDENT_WIDTH   As Long = 10
    Const EXTRA_WIDTH    As Long = 10
    Const EXTRA_HEIGHT   As Long = 100

    Dim VisibleRowCount  As Long
    Dim NewTreeWidth     As Long
    Dim NewTreeHeight    As Long
    Dim NewFormWidth     As Long
    Dim NewFormHeight    As Long

    VisibleRowCount = CountVisibleRows()

    If VisibleRowCount < 1 Then VisibleRowCount = 1
    NewTreeWidth = LongestName * 10 + TreeDepth * INDENT_WIDTH + EXTRA_WIDTH

    NewTreeHeight = VisibleRowCount * ROW_HEIGHT
    NewFormWidth  = NewTreeWidth  + 50
    NewFormHeight = NewTreeHeight + EXTRA_HEIGHT

    Me.Width  = NewFormWidth  + 100
    Me.Height = NewFormHeight + 100

    ValueTree.Width  = NewTreeWidth
    ValueTree.Height = NewTreeHeight
End Sub

Private Function CountVisibleRows() As Long
    Dim TreeNode As MSComctlLib.Node

    If ValueTree.Nodes.Count = 0 Then Exit Function

    For Each TreeNode In ValueTree.Nodes
        If IsNodeVisible(TreeNode) Then
            CountVisibleRows = CountVisibleRows + 1
        End If
    Next TreeNode
End Function

Private Function IsNodeVisible(ByVal TreeNode As MSComctlLib.Node) As Boolean

    Dim ParentNode As MSComctlLib.Node

    If TreeNode Is ValueTree.Nodes(1) Then
        IsNodeVisible = True
        Exit Function
    End If

    Set ParentNode = TreeNode.Parent

    Do While Not ParentNode Is Nothing
        If Not ParentNode.Expanded Then Exit Function

        If ParentNode Is ValueTree.Nodes(1) Then
            IsNodeVisible = True
            Exit Function
        End If

        Set ParentNode = ParentNode.Parent
    Loop
End Function