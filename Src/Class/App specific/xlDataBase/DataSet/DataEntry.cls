VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "DataEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'============================================================
'====================[Properties]============================
'============================================================

Private mValue As Variant

'============================================================
'====================[Getters/Letters]=======================
'============================================================

Public Property Get Value() As Variant
    Value = mValue
End Property

Public Property Let Value(ByVal Value As Variant)
    mValue = Value
End Property

'============================================================
'====================[Creation]==============================
'============================================================

Public Function Create(ByVal Value As Variant) As DataEntry
    Set Create = New DataEntry
    Create.Value = Value
End Function

Public Function Copy(ByVal Other As DataEntry) As DataEntry
    Set Copy = New DataEntry
    Copy.Value = Other.Value
End Function