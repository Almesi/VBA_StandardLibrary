VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "ProgressModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    Title     As String
    Message   As String
    Progress  As Double
    CanCancel As Boolean
    Cancelled As Boolean
End Type

Private This As TModel

Public Property Let Title(ByVal     Value As String)       : Let This.Title     = Value          : End Property
Public Property Let Message(ByVal   Value As String)       : Let This.Message   = Value          : End Property
Public Property Let Progress(ByVal  Value As Double)       : Let This.Progress  = Value          : End Property
Public Property Let CanCancel(ByVal Value As Boolean)      : Let This.CanCancel = Value          : End Property
Public Property Let Cancelled(ByVal Value As Boolean)      : Let This.Cancelled = Value          : End Property
Public Property Get Title()               As String        : Let Title          = This.Title     : End Property
Public Property Get Message()             As String        : Let Message        = This.Message   : End Property
Public Property Get Progress()            As Double        : Let Progress       = This.Progress  : End Property
Public Property Get CanCancel()           As Boolean       : Let CanCancel      = This.CanCancel : End Property
Public Property Get Cancelled()           As Boolean       : Let Cancelled      = This.Cancelled : End Property
Public Property Get Self()                As ProgressModel : Set Self           = Me             : End Property