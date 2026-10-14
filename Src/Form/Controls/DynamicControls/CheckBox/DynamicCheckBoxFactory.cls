VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicCheckBoxFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFunctions         As FunctionDir
Private pCommon            As CommonControlProperties
Private pValue             As Variant
Private pCaption           As String

Private Sub Class_Initialize()
    Set pFunctions     = Nothing
    Set pCommon        = New CommonControlProperties
    pValue             = False
    pCaption           = vbNullString
End Sub


Public Property Set Functions(ByVal           Value As FunctionDir)                : Set pFunctions         = Value             : End Property
Public Property Set Common(ByVal              Value As CommonControlProperties)    : Set pCommon            = Value             : End Property
Public Property Let Value(ByVal               Value As Variant)                    : Let pValue             = Value             : End Property
Public Property Let Caption(ByVal             Value As String)                     : Let pCaption           = Value             : End Property

Public Property Get Functions()               As FunctionDir                       : Set Functions          = pFunctions        : End Property
Public Property Get Common()                  As CommonControlProperties           : Set Common             = pCommon           : End Property
Public Property Get Value()                   As Variant                           : Let Value              = pValue            : End Property
Public Property Get Caption()                 As String                            : Let Caption            = pCaption          : End Property


Public Function Create() As DynamicCheckBoxFactory
    Set Create = New DynamicCheckBoxFactory
End Function

Public Function CreateControl(ByVal Destination As Object) As DynamicCheckBox
    Dim Control As MSForms.CheckBox
    Set Control = Destination.Controls.Add("Forms.CheckBox.1", Common.Name, True)
    Dim Events As CheckBoxEvents
    Set Events = CheckBoxEvents.Create(Control, pFunctions)
    With Control
        Call Common.Apply(Control)
        .Value           = pValue
        .Caption         = pCaption
    End With
    
    Set CreateControl = DynamicCheckBox.Create(Control, Events)
End Function