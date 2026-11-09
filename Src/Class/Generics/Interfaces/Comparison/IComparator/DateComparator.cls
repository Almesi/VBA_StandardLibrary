VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateComparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IComparator

Private Function IComparator_Compare(ByVal LeftValue As Variant, ByVal RightValue As Variant) As Long
    If VarType(LeftValue) <> vbDate Or VarType(RightValue) <> vbDate Then
        Err.Raise 13, "DateComparator", "Both operands must be Date."
    End If
    If LeftValue < RightValue Then
        IComparator_Compare = -1
    ElseIf LeftValue > RightValue Then
        IComparator_Compare = 1
    Else
        IComparator_Compare = 0
    End If

End Function