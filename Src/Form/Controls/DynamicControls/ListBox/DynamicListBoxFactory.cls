VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicListBoxFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFunctions         As FunctionDir
Private pCommon            As CommonControlProperties
Private pValue             As Variant
Private pMultiSelect       As fmMultiSelect
Private pListStyle         As fmListStyle
Private pColumnCount       As Long
Private pBoundColumn       As Long
Private pListRows          As Long
Private pColumnWidths      As String
Private pIntegralHeight    As Boolean

Private Sub Class_Initialize()
    Set pFunctions     = Nothing
    Set pCommon        = New CommonControlProperties
    pValue             = vbNullString
    pMultiSelect       = fmMultiSelectSingle
    pListStyle         = fmListStylePlain
    pColumnCount       = 1
    pBoundColumn       = 1
    pListRows          = 8
    pColumnWidths      = vbNullString
    pIntegralHeight    = True
End Sub

Public Property Set Functions(ByVal Value As FunctionDir)               : Set pFunctions         = Value             : End Property
Public Property Set Common(ByVal Value As CommonControlProperties)      : Set pCommon            = Value             : End Property
Public Property Let Value(ByVal Value As Variant)                       : Let pValue             = Value             : End Property
Public Property Let MultiSelect(ByVal Value As fmMultiSelect)           : Let pMultiSelect       = Value             : End Property
Public Property Let ListStyle(ByVal Value As fmListStyle)               : Let pListStyle         = Value             : End Property
Public Property Let ColumnCount(ByVal Value As Long)                    : Let pColumnCount       = Value             : End Property
Public Property Let BoundColumn(ByVal Value As Long)                    : Let pBoundColumn       = Value             : End Property
Public Property Let ListRows(ByVal Value As Long)                       : Let pListRows          = Value             : End Property
Public Property Let ColumnWidths(ByVal Value As String)                 : Let pColumnWidths      = Value             : End Property
Public Property Let IntegralHeight(ByVal Value As Boolean)              : Let pIntegralHeight    = Value             : End Property

Public Property Get Functions()               As FunctionDir            : Set Functions          = pFunctions        : End Property
Public Property Get Common()                  As CommonControlProperties: Set Common             = pCommon           : End Property
Public Property Get Value()                   As Variant                : Let Value              = pValue            : End Property
Public Property Get MultiSelect()             As fmMultiSelect          : Let MultiSelect        = pMultiSelect      : End Property
Public Property Get ListStyle()               As fmListStyle            : Let ListStyle          = pListStyle        : End Property
Public Property Get ColumnCount()             As Long                   : Let ColumnCount        = pColumnCount      : End Property
Public Property Get BoundColumn()             As Long                   : Let BoundColumn        = pBoundColumn      : End Property
Public Property Get ListRows()                As Long                   : Let ListRows           = pListRows         : End Property
Public Property Get ColumnWidths()            As String                 : Let ColumnWidths       = pColumnWidths     : End Property
Public Property Get IntegralHeight()          As Boolean                : Let IntegralHeight     = pIntegralHeight   : End Property


Public Function Create() As DynamicListBoxFactory
    Set Create = New DynamicListBoxFactory
End Function

Public Function CreateControl(ByVal Destination As Object) As DynamicListBox
    Dim Control As MSForms.ListBox
    Set Control = Destination.Controls.Add("Forms.ListBox.1", Common.Name, True)
    Dim Events As ListBoxEvents
    Set Events = ListBoxEvents.Create(Control, pFunctions)
    With Control
        Call Common.Apply(Control)
        .Value            = pValue
        .MultiSelect      = pMultiSelect
        .ListStyle        = pListStyle
        .ColumnCount      = pColumnCount
        .BoundColumn      = pBoundColumn
        .ListRows         = pListRows
        .ColumnWidths     = pColumnWidths
        .IntegralHeight   = pIntegralHeight
    End With
    
    Set CreateControl = DynamicListBox.Create(Control, Events)
End Function