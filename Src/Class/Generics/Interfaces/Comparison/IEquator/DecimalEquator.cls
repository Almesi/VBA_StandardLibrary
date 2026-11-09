VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecimalEquator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IEquator
Private Function IEquator_AreEqual(ByVal LeftValue As Variant, ByVal RightValue As Variant) As Boolean
    If VarType(LeftValue) <> vbDecimal Or VarType(RightValue) <> vbDecimal Then
        Err.Raise 13, "DecimalEquator", "Both operands must be Decimal Variants."
    End If
    IEquator_AreEqual = (LeftValue = RightValue)

End Function