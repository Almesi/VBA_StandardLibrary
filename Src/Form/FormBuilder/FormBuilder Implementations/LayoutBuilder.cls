VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IFormBuilder

Private mBuilder As IFormBuilder

Private mLayout    As LayoutType
Private mPadWidth  As Single
Private mPadHeight As Single

Public Enum LayoutType
    LayoutRow = 1
    LayoutColumn = 2
End Enum

Public Function Create(ByVal Builder As IFormBuilder, ByVal Layout As LayoutType, Optional ByVal PadWidth As Single = 0, Optional ByVal PadHeight As Single = 0) As LayoutBuilder
    Set Create = New LayoutBuilder
    Call Create.Initialize(Builder, Layout, PadWidth, PadHeight)
End Function

Friend Sub Initialize(ByVal Builder As IFormBuilder, ByVal Layout As LayoutType, ByVal PadWidth As Single, ByVal PadHeight As Single)
    Set mBuilder   = Builder
    Let mLayout    = Layout
    Let mPadWidth  = PadWidth
    Let mPadHeight = PadHeight
End Sub

Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Dim ChildScope As FormBuilderScope

    Set ChildScope = FormBuilderScope.Create(Context.Form, Scope)
    Call Scope.AddScope(ChildScope)
    Call mBuilder.Build(Context, ChildScope)
    Call ApplyLayout(Context.Form.LifeTime, ChildScope)
End Sub

Private Sub ApplyLayout(ByVal LifeTime As Collection, ByVal Scope As FormBuilderScope)
    Select Case mLayout
        Case LayoutRow      : Call ApplyRow(LifeTime, Scope)
        Case LayoutColumn   : Call ApplyColumn(LifeTime, Scope)
    End Select
End Sub

'====================================================================
' ROW
'
' A row arranges its direct children horizontally.
'
' A direct child can be:
'
'   1. A control index
'   2. Another Scope
'
' If it is a Scope, the entire Scope is treated as one group.
'====================================================================

Private Function ApplyRow(ByVal LifeTime As Collection, ByVal Scope As FormBuilderScope) As Single
    Dim i           As Long
    Dim Control     As Object
    Dim CurrentLeft As Single
    Dim ChildLeft   As Single
    Dim ChildWidth  As Single

    CurrentLeft = GetLeft(Scope, LifeTime)

    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            ChildLeft  = GetLeft(Scope.Node(i), LifeTime)
            ChildWidth = GetWidth(Scope.Node(i), LifeTime)

            Call MoveScope(LifeTime, Scope.Node(i), CurrentLeft - ChildLeft, 0)
            CurrentLeft = CurrentLeft + ChildWidth + mPadWidth
        Else
            Set Control = LifeTime.Item(Scope.Node(i)).Control
            Control.Left = CurrentLeft
            CurrentLeft = CurrentLeft + Control.Width + mPadWidth
        End If
    Next i
    If Scope.Count > 0 Then
        CurrentLeft = CurrentLeft - mPadWidth
    End If
    ApplyRow = CurrentLeft
End Function


'====================================================================
' COLUMN
'
' A column arranges its direct children vertically.
'
' A direct child can be:
'
'   1. A control index
'   2. Another Scope
'
' If it is a Scope, the entire Scope is treated as one group.
'====================================================================

Private Function ApplyColumn(ByVal LifeTime As Collection, ByVal Scope As FormBuilderScope) As Single
    Dim i           As Long
    Dim Control     As Object
    Dim CurrentTop  As Single
    Dim ChildTop    As Single
    Dim ChildHeight As Single

    CurrentTop = GetTop(Scope, LifeTime)
    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            ChildTop    = GetTop(Scope.Node(i), LifeTime)
            ChildHeight = GetHeight(Scope.Node(i), LifeTime)

            Call MoveScope(LifeTime, Scope.Node(i), 0, CurrentTop - ChildTop)
            CurrentTop = CurrentTop + ChildHeight + mPadHeight
        Else
            Set Control = LifeTime.Item(Scope.Node(i)).Control
            Control.Top = CurrentTop
            CurrentTop = CurrentTop + Control.Height + mPadHeight
        End If
    Next i

    If Scope.Count > 0 Then
        CurrentTop = CurrentTop - mPadHeight
    End If
    ApplyColumn = CurrentTop
End Function


Private Function IsScope(ByVal Node As Variant) As Boolean
    IsScope = IsObject(Node)
End Function


Private Sub MoveScope(ByVal LifeTime As Collection, ByVal Scope As FormBuilderScope, ByVal DeltaLeft As Single, ByVal DeltaTop As Single)
    Dim i As Long
    Dim Control As Object
    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            Call MoveScope(LifeTime, Scope.Node(i), DeltaLeft, DeltaTop)
        Else
            Set Control = LifeTime.Item(Scope.Node(i)).Control
            Control.Left = Control.Left + DeltaLeft
            Control.Top  = Control.Top  + DeltaTop
        End If
    Next i
End Sub


Private Function GetLeft(ByVal Scope As FormBuilderScope, ByVal LifeTime As Collection) As Single
    Dim i     As Long
    Dim Value As Single
    Dim Found As Boolean

    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            Value = GetLeft(Scope.Node(i), LifeTime)
        Else
            Value = LifeTime.Item(Scope.Node(i)).Control.Left
        End If

        If Not Found Then
            GetLeft = Value
            Found = True
        ElseIf Value < GetLeft Then
            GetLeft = Value
        End If
    Next i
End Function


'====================================================================
' GET TOP
'
' Finds the top-most control in the Scope.
'====================================================================

Private Function GetTop(ByVal Scope As FormBuilderScope, ByVal LifeTime As Collection) As Single
    Dim i     As Long
    Dim Value As Single
    Dim Found As Boolean

    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            Value = GetTop(Scope.Node(i), LifeTime)
        Else
            Value = LifeTime.Item(Scope.Node(i)).Control.Top
        End If
        If Not Found Then
            GetTop = Value
            Found = True
        ElseIf Value < GetTop Then
            GetTop = Value
        End If
    Next i
End Function


Private Function GetWidth(ByVal Scope As FormBuilderScope, ByVal LifeTime As Collection) As Single
    Dim i     As Long
    Dim Left  As Single
    Dim Right As Single
    Dim Value As Single

    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            Left = GetLeft(Scope.Node(i), LifeTime)
            Value = Left + GetWidth(Scope.Node(i), LifeTime)
        Else
            Left = LifeTime.Item(Scope.Node(i)).Control.Left
            Value = Left + LifeTime.Item(Scope.Node(i)).Control.Width
        End If
        If Value > Right Then
            Right = Value
        End If
    Next i
    GetWidth = Right - GetLeft(Scope, LifeTime)
End Function


Private Function GetHeight(ByVal Scope As FormBuilderScope, ByVal LifeTime As Collection) As Single
    Dim i      As Long
    Dim Top    As Single
    Dim Bottom As Single
    Dim Value  As Single

    For i = 1 To Scope.Count
        If IsScope(Scope.Node(i)) Then
            Top = GetTop(Scope.Node(i), LifeTime)
            Value = Top + GetHeight(Scope.Node(i), LifeTime)
        Else
            Top = LifeTime.Item(Scope.Node(i)).Control.Top
            Value = Top + LifeTime.Item(Scope.Node(i)).Control.Height
        End If
        If Value > Bottom Then
            Bottom = Value
        End If
    Next i
    GetHeight = Bottom - GetTop(Scope, LifeTime)
End Function