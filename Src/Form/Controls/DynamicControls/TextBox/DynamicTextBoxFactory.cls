VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicTextBoxFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFunctions         As FunctionDir
Private pCommon            As CommonControlProperties
Private pValue             As Variant
Private pText              As String
Private pMultiLine         As Boolean
Private pPasswordChar      As String
Private pScrollBars        As fmScrollBars
Private pEnterKeyBehavior  As Boolean
Private pWordWrap          As Boolean
Private pAutoSize          As Boolean
Private pTextAlign         As fmTextAlign
Private pLocked            As Boolean
Private pMaxLength         As Long

Private Sub Class_Initialize()
    Set pFunctions     = Nothing
    Set pCommon        = New CommonControlProperties
    pValue             = vbNullString
    pText              = vbNullString
    pMultiLine         = False
    pPasswordChar      = vbNullString
    pScrollBars        = fmScrollBarsNone
    pEnterKeyBehavior  = False
    pWordWrap          = True
    pAutoSize          = False
    pTextAlign         = fmTextAlignLeft
    pLocked            = False
    pMaxLength         = 0
End Sub

Public Property Set Functions(ByVal Value As FunctionDir)               : Set pFunctions         = Value             : End Property
Public Property Set Common(ByVal Value As CommonControlProperties)      : Set pCommon            = Value             : End Property
Public Property Let Value(ByVal Value As Variant)                       : Let pValue             = Value             : End Property
Public Property Let Text(ByVal Value As String)                         : Let pText              = Value             : End Property
Public Property Let MultiLine(ByVal Value As Boolean)                   : Let pMultiLine         = Value             : End Property
Public Property Let PasswordChar(ByVal Value As String)                 : Let pPasswordChar      = Value             : End Property
Public Property Let ScrollBars(ByVal Value As fmScrollBars)             : Let pScrollBars        = Value             : End Property
Public Property Let EnterKeyBehavior(ByVal Value As Boolean)            : Let pEnterKeyBehavior  = Value             : End Property
Public Property Let WordWrap(ByVal Value As Boolean)                    : Let pWordWrap          = Value             : End Property
Public Property Let AutoSize(ByVal Value As Boolean)                    : Let pAutoSize          = Value             : End Property
Public Property Let TextAlign(ByVal Value As fmTextAlign)               : Let pTextAlign         = Value             : End Property
Public Property Let Locked(ByVal Value As Boolean)                      : Let pLocked            = Value             : End Property
Public Property Let MaxLength(ByVal Value As Long)                      : Let pMaxLength         = Value             : End Property

Public Property Get Functions()               As FunctionDir            : Set Functions          = pFunctions        : End Property
Public Property Get Common()                  As CommonControlProperties: Set Common             = pCommon           : End Property
Public Property Get Value()                   As Variant                : Let Value              = pValue            : End Property
Public Property Get Text()                    As String                 : Let Text               = pText             : End Property
Public Property Get MultiLine()               As Boolean                : Let MultiLine          = pMultiLine        : End Property
Public Property Get PasswordChar()            As String                 : Let PasswordChar       = pPasswordChar     : End Property
Public Property Get ScrollBars()              As fmScrollBars           : Let ScrollBars         = pScrollBars       : End Property
Public Property Get EnterKeyBehavior()        As Boolean                : Let EnterKeyBehavior   = pEnterKeyBehavior : End Property
Public Property Get WordWrap()                As Boolean                : Let WordWrap           = pWordWrap         : End Property
Public Property Get AutoSize()                As Boolean                : Let AutoSize           = pAutoSize         : End Property
Public Property Get TextAlign()               As fmTextAlign            : Let TextAlign          = pTextAlign        : End Property
Public Property Get Locked()                  As Boolean                : Let Locked             = pLocked           : End Property
Public Property Get MaxLength()              As Long                    : Let MaxLength          = pMaxLength        : End Property


Public Function Create() As DynamicTextBoxFactory
    Set Create = New DynamicTextBoxFactory
End Function

Public Function CreateControl(ByVal Destination As Object) As DynamicTextBox
    Dim Control As MSForms.TextBox
    Set Control = Destination.Controls.Add("Forms.TextBox.1", Common.Name, True)
    Dim Events As TextBoxEvents
    Set Events = TextBoxEvents.Create(Control, pFunctions)
    With Control
        Call Common.Apply(Control)
        .Value            = pValue
        .Text             = pText
        .MultiLine        = pMultiLine
        .PasswordChar     = pPasswordChar
        .ScrollBars       = pScrollBars
        .EnterKeyBehavior = pEnterKeyBehavior
        .WordWrap         = pWordWrap
        .AutoSize         = pAutoSize
        .TextAlign        = pTextAlign
        .Locked           = pLocked
        .MaxLength        = pMaxLength
    End With
    
    Set CreateControl = DynamicTextBox.Create(Control, Events)
End Function