VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormLayoutBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IFormBuilder

Private mBuilder      As IFormBuilder
Private mAutoSize     As Boolean
Private mMaxWidth     As Single
Private mMaxHeight    As Single
Private mScrollBars   As Long
Private mScrollWidth  As Single
Private mScrollHeight As Single
Private mMargin       As Single
Private mFormWidth    As Single
Private mFormHeight   As Single

Private Const DEFAULT_MAX_WIDTH      As Single = 600
Private Const DEFAULT_MAX_HEIGHT     As Single = 450
Private Const DEFAULT_FORM_WIDTH     As Single = 600
Private Const DEFAULT_FORM_HEIGHT    As Single = 450
Private Const DEFAULT_MARGIN         As Single = 100
Private Const SCROLL_BARS_NONE       As Long   = 0
Private Const SCROLL_BARS_HORIZONTAL As Long   = 1
Private Const SCROLL_BARS_VERTICAL   As Long   = 2
Private Const SCROLL_BARS_BOTH       As Long   = 3

Private Sub Class_Initialize()
    Set mBuilder = Nothing

    mAutoSize     = True
    mMaxWidth     = DEFAULT_MAX_WIDTH
    mMaxHeight    = DEFAULT_MAX_HEIGHT
    mScrollBars   = SCROLL_BARS_BOTH
    mScrollWidth  = DEFAULT_MAX_WIDTH
    mScrollHeight = DEFAULT_MAX_HEIGHT
    mMargin       = DEFAULT_MARGIN
    mFormWidth    = DEFAULT_FORM_WIDTH
    mFormHeight   = DEFAULT_FORM_HEIGHT
End Sub

Public Function Create(ByVal Builder As IFormBuilder) As FormLayoutBuilder
    Dim Result As FormLayoutBuilder
    Set Result = New FormLayoutBuilder
    Set Result.Builder = Builder
    Set Create = Result
End Function

Public Property Set Builder(ByVal      Value As IFormBuilder) : Set mBuilder      = Value         : End Property
Public Property Let AutoSize(ByVal     Value As Boolean)      : Let mAutoSize     = Value         : End Property
Public Property Let MaxWidth(ByVal     Value As Single)       : Let mMaxWidth     = Value         : End Property
Public Property Let MaxHeight(ByVal    Value As Single)       : Let mMaxHeight    = Value         : End Property
Public Property Let ScrollBars(ByVal   Value As Long)         : Let mScrollBars   = Value         : End Property
Public Property Let ScrollWidth(ByVal  Value As Single)       : Let mScrollWidth  = Value         : End Property
Public Property Let ScrollHeight(ByVal Value As Single)       : Let mScrollHeight = Value         : End Property
Public Property Let Margin(ByVal       Value As Single)       : Let mMargin       = Value         : End Property
Public Property Let FormWidth(ByVal    Value As Single)       : Let mFormWidth    = Value         : End Property
Public Property Let FormHeight(ByVal   Value As Single)       : Let mFormHeight   = Value         : End Property
Public Property Get Builder()                As IFormBuilder  : Set Builder       = mBuilder      : End Property
Public Property Get AutoSize()               As Boolean       : Let AutoSize      = mAutoSize     : End Property
Public Property Get MaxWidth()               As Single        : Let MaxWidth      = mMaxWidth     : End Property
Public Property Get MaxHeight()              As Single        : Let MaxHeight     = mMaxHeight    : End Property
Public Property Get ScrollBars()             As Long          : Let ScrollBars    = mScrollBars   : End Property
Public Property Get ScrollWidth()            As Single        : Let ScrollWidth   = mScrollWidth  : End Property
Public Property Get ScrollHeight()           As Single        : Let ScrollHeight  = mScrollHeight : End Property
Public Property Get Margin()                 As Single        : Let Margin        = mMargin       : End Property
Public Property Get FormWidth()              As Single        : Let FormWidth     = mFormWidth    : End Property
Public Property Get FormHeight()             As Single        : Let FormHeight    = mFormHeight   : End Property

Private Sub IFormBuilder_Build(ByVal Context As FormBuilderContext, ByVal Scope As FormBuilderScope)
    Dim TargetForm      As Object
    Dim ContentWidth    As Single
    Dim ContentHeight   As Single
    Dim DesiredWidth    As Single
    Dim DesiredHeight   As Single
    Dim ViewWidth       As Single
    Dim ViewHeight      As Single
    Dim NeedHorizontal  As Boolean
    Dim NeedVertical    As Boolean
    Dim ScrollBarWidth  As Single
    Dim ScrollBarHeight As Single

    Call mBuilder.Build(Context, Scope)

    Set TargetForm = Context.Form.Form
    ContentWidth   = GetTotalControlWidth(TargetForm)  + mMargin
    ContentHeight  = GetTotalControlHeight(TargetForm) + mMargin

    If mAutoSize Then
        If ContentWidth  > mMaxWidth  And mMaxWidth  > 0 Then DesiredWidth  = mMaxWidth  Else DesiredWidth  = ContentWidth
        If ContentHeight > mMaxHeight And mMaxHeight > 0 Then DesiredHeight = mMaxHeight Else DesiredHeight = ContentHeight
    Else
        DesiredWidth  = mFormWidth
        DesiredHeight = mFormHeight
    End If

    TargetForm.Width  = DesiredWidth
    TargetForm.Height = DesiredHeight

    NeedHorizontal = DesiredWidth  < ContentWidth
    NeedVertical   = DesiredHeight < ContentHeight

    If mAutoSize Then
        Select Case True
            Case NeedHorizontal And NeedVertical  : TargetForm.ScrollBars = SCROLL_BARS_BOTH
            Case NeedHorizontal                   : TargetForm.ScrollBars = SCROLL_BARS_HORIZONTAL
            Case NeedVertical                     : TargetForm.ScrollBars = SCROLL_BARS_VERTICAL
            Case Else                             : TargetForm.ScrollBars = SCROLL_BARS_NONE
        End Select
        If NeedHorizontal Then TargetForm.ScrollWidth  = ContentWidth
        If NeedVertical   Then TargetForm.ScrollHeight = ContentHeight
    Else
        TargetForm.ScrollBars   = mScrollBars
        TargetForm.ScrollWidth  = mScrollWidth
        TargetForm.ScrollHeight = mScrollHeight
    End If
    TargetForm.ScrollLeft = 0
    TargetForm.ScrollTop  = 0
End Sub

Private Function GetTotalControlWidth(ByVal TargetForm As Object) As Single
    Dim Control As Object
    Dim Max     As Single
    Dim Current As Single

    For Each Control In TargetForm.Controls
        Current = Control.Left + Control.Width
        If Current > Max Then Max = Current
    Next Control
    GetTotalControlWidth = Max
End Function

Private Function GetTotalControlHeight(ByVal TargetForm As Object) As Single
    Dim Control As Object
    Dim Max     As Single
    Dim Current As Single

    For Each Control In TargetForm.Controls
        Current = Control.Top + Control.Height
        If Current > Max Then Max = Current
    Next Control
    GetTotalControlHeight = Max
End Function