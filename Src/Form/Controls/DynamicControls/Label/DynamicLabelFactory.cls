VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicLabelFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFunctions         As FunctionDir
Private pCommon            As CommonControlProperties
Private pCaption           As String
Private pAutoSize          As Boolean
Private pWordWrap          As Boolean
Private pTextAlign         As fmTextAlign
Private pBackStyle         As fmBackStyle
Private pBorderStyle       As fmBorderStyle
Private pSpecialEffect     As fmSpecialEffect

Private Sub Class_Initialize()
    Set pFunctions     = Nothing
    Set pCommon        = New CommonControlProperties
    pCaption           = vbNullString
    pAutoSize          = False
    pWordWrap          = True
    pTextAlign         = fmTextAlignLeft
    pBackStyle         = fmBackStyleOpaque
    pBorderStyle       = fmBorderStyleNone
    pSpecialEffect     = fmSpecialEffectFlat
End Sub

Public Property Set Functions(ByVal Value As FunctionDir)               : Set pFunctions         = Value             : End Property
Public Property Set Common(ByVal Value As CommonControlProperties)      : Set pCommon            = Value             : End Property
Public Property Let Caption(ByVal Value As String)                      : Let pCaption           = Value             : End Property
Public Property Let AutoSize(ByVal Value As Boolean)                    : Let pAutoSize          = Value             : End Property
Public Property Let WordWrap(ByVal Value As Boolean)                    : Let pWordWrap          = Value             : End Property
Public Property Let TextAlign(ByVal Value As fmTextAlign)               : Let pTextAlign         = Value             : End Property
Public Property Let BackStyle(ByVal Value As fmBackStyle)               : Let pBackStyle         = Value             : End Property
Public Property Let BorderStyle(ByVal Value As fmBorderStyle)           : Let pBorderStyle       = Value             : End Property
Public Property Let SpecialEffect(ByVal Value As fmSpecialEffect)       : Let pSpecialEffect     = Value             : End Property

Public Property Get Functions()               As FunctionDir            : Set Functions          = pFunctions        : End Property
Public Property Get Common()                  As CommonControlProperties: Set Common             = pCommon           : End Property
Public Property Get Caption()                 As String                 : Let Caption            = pCaption          : End Property
Public Property Get AutoSize()                As Boolean                : Let AutoSize           = pAutoSize         : End Property
Public Property Get WordWrap()                As Boolean                : Let WordWrap           = pWordWrap         : End Property
Public Property Get TextAlign()               As fmTextAlign            : Let TextAlign          = pTextAlign        : End Property
Public Property Get BackStyle()               As fmBackStyle            : Let BackStyle          = pBackStyle        : End Property
Public Property Get BorderStyle()             As fmBorderStyle          : Let BorderStyle        = pBorderStyle      : End Property
Public Property Get SpecialEffect()           As fmSpecialEffect        : Let SpecialEffect      = pSpecialEffect    : End Property


Public Function Create() As DynamicLabelFactory
    Set Create = New DynamicLabelFactory
End Function

Public Function CreateControl(ByVal Destination As Object) As DynamicLabel
    Dim Control As MSForms.Label
    Set Control = Destination.Controls.Add("Forms.Label.1", Common.Name, True)
    Dim Events As LabelEvents
    Set Events = LabelEvents.Create(Control, pFunctions)
    With Control
        Call Common.Apply(Control)
        .Caption          = pCaption
        .AutoSize         = pAutoSize
        .WordWrap         = pWordWrap
        .TextAlign        = pTextAlign
        .BackStyle        = pBackStyle
        .BorderStyle      = pBorderStyle
        .SpecialEffect    = pSpecialEffect
    End With
    
    Set CreateControl = DynamicLabel.Create(Control, Events)
End Function