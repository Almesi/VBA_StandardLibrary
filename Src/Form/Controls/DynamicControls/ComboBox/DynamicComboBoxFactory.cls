VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "DynamicComboBoxFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private pFunctions         As FunctionDir
Private pCommon            As CommonControlProperties
Private pValue             As Variant
Private pText              As String
Private pStyle             As fmStyle
Private pMatchRequired     As Boolean
Private pMatchEntry        As fmMatchEntry
Private pColumnCount       As Long
Private pBoundColumn       As Long
Private pListRows          As Long
Private pListWidth         As Double
Private pColumnWidths      As String

Private Sub Class_Initialize()
    Set pFunctions     = Nothing
    Set pCommon        = New CommonControlProperties
    pValue             = vbNullString
    pText              = vbNullString
    pStyle             = fmStyleDropDownCombo
    pMatchRequired     = False
    pMatchEntry        = fmMatchEntryFirstLetter
    pColumnCount       = 1
    pBoundColumn       = 1
    pListRows          = 8
    pListWidth         = 0
    pColumnWidths      = vbNullString
End Sub

Public Property Set Functions(ByVal Value As FunctionDir)               : Set pFunctions         = Value             : End Property
Public Property Set Common(ByVal Value As CommonControlProperties)      : Set pCommon            = Value             : End Property
Public Property Let Value(ByVal Value As Variant)                       : Let pValue             = Value             : End Property
Public Property Let Text(ByVal Value As String)                         : Let pText              = Value             : End Property
Public Property Let Style(ByVal Value As fmStyle)                       : Let pStyle             = Value             : End Property
Public Property Let MatchRequired(ByVal Value As Boolean)               : Let pMatchRequired     = Value             : End Property
Public Property Let MatchEntry(ByVal Value As fmMatchEntry)             : Let pMatchEntry        = Value             : End Property
Public Property Let ColumnCount(ByVal Value As Long)                    : Let pColumnCount       = Value             : End Property
Public Property Let BoundColumn(ByVal Value As Long)                    : Let pBoundColumn       = Value             : End Property
Public Property Let ListRows(ByVal Value As Long)                       : Let pListRows          = Value             : End Property
Public Property Let ListWidth(ByVal Value As Double)                    : Let pListWidth         = Value             : End Property
Public Property Let ColumnWidths(ByVal Value As String)                 : Let pColumnWidths      = Value             : End Property

Public Property Get Functions()               As FunctionDir            : Set Functions          = pFunctions        : End Property
Public Property Get Common()                  As CommonControlProperties: Set Common             = pCommon           : End Property
Public Property Get Value()                   As Variant                : Let Value              = pValue            : End Property
Public Property Get Text()                    As String                 : Let Text               = pText             : End Property
Public Property Get Style()                   As fmStyle                : Let Style              = pStyle            : End Property
Public Property Get MatchRequired()           As Boolean                : Let MatchRequired      = pMatchRequired    : End Property
Public Property Get MatchEntry()              As fmMatchEntry           : Let MatchEntry         = pMatchEntry       : End Property
Public Property Get ColumnCount()             As Long                   : Let ColumnCount        = pColumnCount      : End Property
Public Property Get BoundColumn()             As Long                   : Let BoundColumn        = pBoundColumn      : End Property
Public Property Get ListRows()                As Long                   : Let ListRows           = pListRows         : End Property
Public Property Get ListWidth()               As Double                 : Let ListWidth          = pListWidth        : End Property
Public Property Get ColumnWidths()            As String                 : Let ColumnWidths       = pColumnWidths     : End Property


Public Function Create() As DynamicComboBoxFactory
    Set Create = New DynamicComboBoxFactory
End Function

Public Function CreateControl(ByVal Destination As Object) As DynamicComboBox
    Dim Control As MSForms.ComboBox
    Set Control = Destination.Controls.Add("Forms.ComboBox.1", Common.Name, True)
    Dim Events As ComboBoxEvents
    Set Events = ComboBoxEvents.Create(Control, pFunctions)
    With Control
        Call Common.Apply(Control)
        .Value            = pValue
        .Text             = pText
        .Style            = pStyle
        .MatchRequired    = pMatchRequired
        .MatchEntry       = pMatchEntry
        .ColumnCount      = pColumnCount
        .BoundColumn      = pBoundColumn
        .ListRows         = pListRows
        .ListWidth        = pListWidth
        .ColumnWidths     = pColumnWidths
    End With
    
    Set CreateControl = DynamicComboBox.Create(Control, Events)
End Function