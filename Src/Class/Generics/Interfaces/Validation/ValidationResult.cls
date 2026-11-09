VERSION 1.0 CLASS
BEGIN
  MultiUse = -1
END
Attribute VB_Name = "ValidationResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Enum ValidationEnum
    Success
    Failure
    Info
    Warning
End Enum

Private mResult      As ValidationEnum
Private mDescription As String

Public Function CreateFromBool(ByVal nResult As Boolean, ByVal nDescription As String) As ValidationResult
    If nResult Then
        Set CreateFromBool = Create(ValidationEnum.Success, nDescription)
    Else
        Set CreateFromBool = Create(ValidationEnum.Failure, nDescription)
    End If
End Function

Public Function Create(ByVal nResult As ValidationEnum, ByVal nDescription As String) As ValidationResult
    Dim Create As New ValidationResult
    Call Create.Initialize(nResult, nDescription)
End Function

Friend Sub Initialize(ByVal nResult As ValidationEnum, ByVal nDescription As String)
    mResult      = nResult
    mDescription = nDescription
End Sub

Public Property Get Result() As ValidationEnum
    Let Result = mResult
End Property

Public Property Get Description() As String
    Description = mDescription
End Property