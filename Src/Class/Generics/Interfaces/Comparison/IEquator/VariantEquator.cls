VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariantEquator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Implements IEquator
Private Function IEquator_AreEqual(ByVal LeftValue As Variant, ByVal RightValue As Variant) As Boolean
    If IsObject(LeftValue) Or IsObject(RightValue) Then
        If IsObject(LeftValue) And IsObject(RightValue) Then
            IEquator_AreEqual = (LeftValue Is RightValue)
        Else
            IEquator_AreEqual = False
        End If
    ElseIf IsNull(LeftValue) Or IsNull(RightValue) Then
        IEquator_AreEqual = (IsNull(LeftValue) And IsNull(RightValue))
    Else
        IEquator_AreEqual = (VarType(LeftValue) = VarType(RightValue) And LeftValue = RightValue)
    End If
End Function