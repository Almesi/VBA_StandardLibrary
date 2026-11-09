VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BooleanComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Implements IComparator

Private Function IComparator_Compare(ByVal LeftValue As Variant, ByVal RightValue As Variant) As Long
    If VarType(LeftValue) <> vbBoolean Or VarType(RightValue) <> vbBoolean Then
        Err.Raise 13, "BooleanComparator", "Both operands must be Boolean."
    End If
    If CBool(LeftValue) = CBool(RightValue) Then
        IComparator_Compare = 0
    ElseIf CBool(LeftValue) Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = -1
    End If

End Function