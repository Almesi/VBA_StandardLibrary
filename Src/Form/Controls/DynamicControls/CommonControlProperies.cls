VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "CommonControlProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pName              As String
Private pLeft              As Single
Private pTop               As Single
Private pWidth             As Single
Private pHeight            As Single
Private pVisible           As Boolean
Private pEnabled           As Boolean
Private pTabIndex          As Integer
Private pControlTipText    As String
Private pTag               As String
Private pBackColor         As Long
Private pForeColor         As Long
Private pFontName          As String
Private pFontSize          As Single
Private pFontBold          As Boolean
Private pFontItalic        As Boolean
Private pFontUnderline     As Boolean

Private Sub Class_Initialize()
    pName              = vbNullString
    pLeft              = 0
    pTop               = 0
    pWidth             = 100
    pHeight            = 30
    pVisible           = True
    pEnabled           = True
    pTabIndex          = 0
    pControlTipText    = vbNullString
    pTag               = vbNullString
    pBackColor         = vbWhite
    pForeColor         = vbBlack
    pFontName          = "Tahoma"
    pFontSize          = 8
    pFontBold          = False
    pFontItalic        = False
    pFontUnderline     = False
End Sub


Public Property Let Name(ByVal                Value As String)                     : Let pName              = Value             : End Property
Public Property Let Left(ByVal                Value As Single)                     : Let pLeft              = Value             : End Property
Public Property Let Top(ByVal                 Value As Single)                     : Let pTop               = Value             : End Property
Public Property Let Width(ByVal               Value As Single)                     : Let pWidth             = Value             : End Property
Public Property Let Height(ByVal              Value As Single)                     : Let pHeight            = Value             : End Property
Public Property Let Visible(ByVal             Value As Boolean)                    : Let pVisible           = Value             : End Property
Public Property Let Enabled(ByVal             Value As Boolean)                    : Let pEnabled           = Value             : End Property
Public Property Let TabIndex(ByVal            Value As Integer)                    : Let pTabIndex          = Value             : End Property
Public Property Let ControlTipText(ByVal      Value As String)                     : Let pControlTipText    = Value             : End Property
Public Property Let Tag(ByVal                 Value As String)                     : Let pTag               = Value             : End Property
Public Property Let BackColor(ByVal           Value As Long)                       : Let pBackColor         = Value             : End Property
Public Property Let ForeColor(ByVal           Value As Long)                       : Let pForeColor         = Value             : End Property
Public Property Let FontName(ByVal            Value As String)                     : Let pFontName          = Value             : End Property
Public Property Let FontSize(ByVal            Value As Single)                     : Let pFontSize          = Value             : End Property
Public Property Let FontBold(ByVal            Value As Boolean)                    : Let pFontBold          = Value             : End Property
Public Property Let FontItalic(ByVal          Value As Boolean)                    : Let pFontItalic        = Value             : End Property
Public Property Let FontUnderline(ByVal       Value As Boolean)                    : Let pFontUnderline     = Value             : End Property

Public Property Get Name()                          As String                      : Let Name               = pName             : End Property
Public Property Get Left()                          As Single                      : Let Left               = pLeft             : End Property
Public Property Get Top()                           As Single                      : Let Top                = pTop              : End Property
Public Property Get Width()                         As Single                      : Let Width              = pWidth            : End Property
Public Property Get Height()                        As Single                      : Let Height             = pHeight           : End Property
Public Property Get Visible()                       As Boolean                     : Let Visible            = pVisible          : End Property
Public Property Get Enabled()                       As Boolean                     : Let Enabled            = pEnabled          : End Property
Public Property Get TabIndex()                      As Integer                     : Let TabIndex           = pTabIndex         : End Property
Public Property Get ControlTipText()                As String                      : Let ControlTipText     = pControlTipText   : End Property
Public Property Get Tag()                           As String                      : Let Tag                = pTag              : End Property
Public Property Get BackColor()                     As Long                        : Let BackColor          = pBackColor        : End Property
Public Property Get ForeColor()                     As Long                        : Let ForeColor          = pForeColor        : End Property
Public Property Get FontName()                      As String                      : Let FontName           = pFontName         : End Property
Public Property Get FontSize()                      As Single                      : Let FontSize           = pFontSize         : End Property
Public Property Get FontBold()                      As Boolean                     : Let FontBold           = pFontBold         : End Property
Public Property Get FontItalic()                    As Boolean                     : Let FontItalic         = pFontItalic       : End Property
Public Property Get FontUnderline()                 As Boolean                     : Let FontUnderline      = pFontUnderline    : End Property


Public Function Create() As CommonControlProperties
    Set Create = New CommonControlProperties
    With Create
        .Left            = pLeft
        .Top             = pTop
        .Width           = pWidth
        .Height          = pHeight
        .Visible         = pVisible
        .Enabled         = pEnabled
        .TabIndex        = pTabIndex
        .ControlTipText  = pControlTipText
        .Tag             = pTag
        .BackColor       = pBackColor
        .ForeColor       = pForeColor
        .FontName        = pFontName
        .FontSize        = pFontSize
        .FontBold        = pFontBold
        .FontItalic      = pFontItalic
        .FontUnderline   = pFontUnderline
    End With
End Function

Public Sub Apply(ByVal Control As Object)
    With Control
        .Left            = pLeft
        .Top             = pTop
        .Width           = pWidth
        .Height          = pHeight
        .Visible         = pVisible
        .Enabled         = pEnabled
        .TabIndex        = pTabIndex
        .ControlTipText  = pControlTipText
        .Tag             = pTag
        .BackColor       = pBackColor
        .ForeColor       = pForeColor
        .Font.Name       = pFontName
        .Font.Size       = pFontSize
        If pFontBold      Then .Font.Bold       = pFontBold
        If pFontItalic    Then .Font.Italic     = pFontItalic
        If pFontUnderline Then .Font.Underline  = pFontUnderline
    End With
End Sub