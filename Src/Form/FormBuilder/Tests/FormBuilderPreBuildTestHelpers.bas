Attribute VB_Name = "FormBuilderPreBuildTestHelpers"

Option Explicit

Public Sub Printt(ByVal Text As String)
    Debug.Print Text
End Sub

Public Sub OKButtonTest1(ByVal View As IView)
    Dim Control As Object
    Set Control = View.Form.Controls(1)

    View.Result = FormReturn.FormReturnOK
    View.Form.Hide

    Call Printt(Control.Name)
End Sub

Public Sub CancelButtonTest1(ByVal View As IView)
    View.Result = FormReturn.FormReturnOK
    View.Form.Hide
End Sub


Public Function MakeFunctionDirSingle(ByVal EventName As String, ByVal Func As IFunction) As FunctionDir
    Dim Result As FunctionDir
    Set Result = FunctionDir.Create()

    Call Result.Add(EventName, Func)
    Set MakeFunctionDirSingle = Result
End Function