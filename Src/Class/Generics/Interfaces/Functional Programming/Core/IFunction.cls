VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True


Option Explicit

Public Function Evaluate(ParamArray Arguments() As Variant) As Variant
Public Function EvaluateArr(ByRef Arguments() As Variant) As Variant

Public Function Run(ParamArray Arguments() As Variant) As IFunction
Public Function RunArr(ByRef Arguments() As Variant) As IFunction