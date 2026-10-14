VERSION 1.0 CLASS
BEGIN
    MultiUse = -1  'True
END
Attribute VB_Name = "ControlValidated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mControl   As IControlValue
Private mValidator As Validator

Public Function Create(ByVal Control As IControlValue, ByVal Validator As Validator) As ControlValidated
    Set Create = New ControlValidated
    Call Create.Initialize(Control, Validator)
End Function

Friend Sub Initialize(ByVal Control As IControlValue, ByVal Validator As Validator)
    Set mControl   = Control
    Set mValidator = Validator
End Sub

Public Function ValidateAndRewrite() As ValidationResult
    Dim NewValue As ValidatedValue
    Dim Result   As ValidationResult
    Set Result = Validate(NewValue)
    If ValidationResult.Result = ValidationEnum.Success Then
        mControl.Value = NewValue.Value
    End If
    Set ValidateAndRewrite = Result
End Function

Public Function Validate(ByVal Output As ValidatedValue) As ValidationResult
    If mValidator Is Nothing Then
        Set Validate = New ValidationResult
        Validate.Result      = ValidationEnum.Info
        Validate.Description = "No Validator provided"
        Exit Function
    End If
    Set Validate = mValidator.Validate(mControl.Value, Output)
End Function