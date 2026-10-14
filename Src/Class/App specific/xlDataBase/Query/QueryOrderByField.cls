VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "QueryOrderByField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mFieldName As String
Private mDescending As Boolean


'===============================================================================
' INITIALIZATION
'===============================================================================

Public Function Create(ByVal FieldName As String, Optional ByVal Descending As Boolean = False) As QueryOrderByField
    Set Create = New QueryOrderByField
    Call Create.Initialize(FieldName, Descending)
End Function


Friend Sub Initialize(ByVal FieldName As String, ByVal Descending As Boolean)
    If Len(Trim(FieldName)) = 0 Then
        Err.Raise vbObjectError + 1430, _
                  "QueryOrderByField.Initialize", _
                  "Field name cannot be empty."
    End If

    Let mFieldName = FieldName
    Let mDescending = Descending
End Sub


'===============================================================================
' PROPERTIES
'===============================================================================

Public Property Get FieldName() As String
    FieldName = mFieldName
End Property


Public Property Get Descending() As Boolean
    Descending = mDescending
End Property