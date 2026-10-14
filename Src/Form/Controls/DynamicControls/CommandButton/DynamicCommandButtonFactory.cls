VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicCommandButtonFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFunctions         As FunctionDir
Private pCommon            As CommonControlProperties
Private pCaption           As String
Private pValue             As Boolean
Private pDefault           As Boolean
Private pCancel            As Boolean

Private Sub Class_Initialize()
    Set pFunctions     = Nothing
    Set pCommon        = New CommonControlProperties
    pCaption           = vbNullString
    pValue             = False
    pDefault           = False
    pCancel            = False
End Sub

Public Property Set Functions(ByVal Value As FunctionDir)               : Set pFunctions         = Value             : End Property
Public Property Set Common(ByVal Value As CommonControlProperties)      : Set pCommon            = Value             : End Property
Public Property Let Caption(ByVal Value As String)                      : Let pCaption           = Value             : End Property
Public Property Let Value(ByVal Value As Boolean)                       : Let pValue             = Value             : End Property
Public Property Let Default(ByVal Value As Boolean)                     : Let pDefault           = Value             : End Property
Public Property Let Cancel(ByVal Value As Boolean)                      : Let pCancel            = Value             : End Property

Public Property Get Functions()               As FunctionDir            : Set Functions          = pFunctions        : End Property
Public Property Get Common()                  As CommonControlProperties: Set Common             = pCommon           : End Property
Public Property Get Caption()                 As String                 : Let Caption            = pCaption          : End Property
Public Property Get Value()                   As Boolean                : Let Value              = pValue            : End Property
Public Property Get Default()                 As Boolean                : Let Default            = pDefault          : End Property
Public Property Get Cancel()                  As Boolean                : Let Cancel             = pCancel           : End Property


Public Function Create() As DynamicCommandButtonFactory
    Set Create = New DynamicCommandButtonFactory
End Function

Public Function CreateControl(ByVal Destination As Object) As DynamicCommandButton
    Dim Control As MSForms.CommandButton
    Set Control = Destination.Controls.Add("Forms.CommandButton.1", Common.Name, True)
    Dim Events As CommandButtonEvents
    Set Events = CommandButtonEvents.Create(Control, pFunctions)
    With Control
        Call Common.Apply(Control)
        .Caption          = pCaption
        .Value            = pValue
        .Default          = pDefault
        .Cancel           = pCancel
    End With
    
    Set CreateControl = DynamicCommandButton.Create(Control, Events)
End Function