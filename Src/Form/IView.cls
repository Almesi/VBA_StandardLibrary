VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
 
Public Function Show(ByVal Model As Object) As FormReturn
Public Property Let Result(ByVal Value As FormReturn)
Public Property Get Form() As Object