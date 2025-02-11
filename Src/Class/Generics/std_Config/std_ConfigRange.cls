VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "std_ConfigRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IConfig

Private p_Destination As Range

Public Event BeforeLet(Indicator As Variant, Value As Variant)
Public Event AfterLet(Indicator As Variant, Value As Variant)
Public Event BeforeGet(Indicator As Variant)
Public Event AfterGet(Indicator As Variant)

Public Function Create(n_Destination As Variant) As std_ConfigRange
    Set Create = New std_ConfigRange
    Create.IConfig_Destination = n_Destination
End Function

Public Property Let IConfig_Destination(n_Destination As Variant)
    Set p_Destination = n_Destination
End Property

Public Property Get IConfig_Destination() As Variant
    Set Destination = p_Destination
End Property

Public Property Let IConfig_Setting(Indicator As Variant, Optional Check As Boolean = False, Optional Index As Long = 0, Value As Variant)  
    Dim Scope As Range
    Dim Cell As Range

    RaiseEvent BeforeLet(Indicator, Value)
    Set Scope = GetScope(Indicator)
    If Not Scope Is Nothing Then
        If Check Then
            If IConfig_Check(Scope, Value) Then
                Scope.Cells(1, 1).Value = Value
            End If
        Else
            If Index <> -1 Then
                Scope.Cells(1, Index + 1).Value = Value
            Else
                For Each Cell In Scope
                    Cell.Value = Value
                Next Cell
            End If
        End If
    End If
    RaiseEvent AfterLet(Indicator, Value)
End Property

Public Property Get IConfig_Setting(Indicator As Variant, Optional Check As Boolean = False, Optional Index As Long = 0) As Variant
    Dim Scope As Range

    RaiseEvent BeforeGet(Indicator)
    Set Scope = GetScope(Indicator)
    If Not Scope Is Nothing Then
        If Check Then
            If IConfig_Check(Scope, Scope.Cells(1, 1)) Then
                Set IConfig_Setting = Scope.Cells(1, 1)
            End If
        Else
            If Index <> -1 Then
                Set IConfig_Setting = Scope.Cells(1, Index + 1)
            Else
                Set IConfig_Setting = Scope
            End If
        End If
    End If
    RaiseEvent AfterGet(Indicator)
End Property

Public Function IConfig_Check(Scope As Variant, Value As Variant) As Boolean
    Dim Cell As Range
    Dim i As Long
    For Each Cell In Scope.Cells
        If Cell.Value = Value Then
            If i <> 0 Then
                IConfig_Check = True
                Exit Function
            End If
        End If
        i = i + 1
    Next Cell
End Function

Public Function GetScope(Indicator As Variant) As Range
    Dim Rng As Range
    Set Rng = p_Destination.Find(Indicator)
    If Not Rng Is Nothing Then
        Set GetScope = Range(Rng.Offset(0, 1), Rng.End(xlToRight))
    End If
End Function