VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_BufferElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private p_Typee As Long
Private p_Count As Long
Private p_Normalized As Byte

Public Property Let Typee(n_Typee As Long)          : p_Typee      = n_Typee     : End Property
Public Property Let Count(n_Count As Long)          : p_Count      = n_Count     : End Property
Public Property Let Normalized(n_Normalized As Byte): p_Normalized = n_Normalized: End Property

Public Property Get Typee() As Long     : Typee      = p_Typee     : End Property
Public Property Get Count() As Long     : Count      = p_Count     : End Property
Public Property Get Normalized() As Byte: Normalized = p_Normalized: End Property