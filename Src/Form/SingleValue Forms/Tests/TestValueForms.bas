Attribute VB_Name = "TestValueForms"

Option Explicit

Public Sub TestAllValueForms()
    Call TestCallValueForm
    Call TestEnumValueForm
    Call TestFreeValueForm
End Sub

Public Function GetDateFromForm() As Date
    Dim View   As IView
    Dim Model  As ValueModel
    Dim Result As FormReturn

    Set Model = New ValueModel
    Set View = DateValueForm.Create()

    Result = View.Show(Model)

    GetDateFromForm = Model.Value
End Function

Public Sub TestCallValueForm()
    Dim View   As IView
    Dim Model  As ValueModel
    Dim Result As FormReturn

    Set Model = New ValueModel
    Set View  = CallValueForm.Create(NamedFunction.Create("GetDateFromForm", False, False))

    Result = View.Show(Model)

    Debug.Print "CallValueForm result: "; Result
    Debug.Print "Model value: "; Model.Value
End Sub

Public Sub TestEnumValueForm()
    Dim View   As IView
    Dim Model  As ValueModel
    Dim Enums  As Collection
    Dim Result As FormReturn

    Set Enums = New Collection
    Enums.Add "Red"
    Enums.Add "Green"
    Enums.Add "Blue"

    Set Model = New ValueModel
    Set View  = EnumValueForm.Create(Enums)

    Result = View.Show(Model)

    Debug.Print "EnumValueForm result: "; Result
    Debug.Print "Model value: "; Model.Value
End Sub

Public Sub TestFreeValueForm()
    Dim View   As IView
    Dim Model  As ValueModel
    Dim Result As FormReturn

    Set Model = New ValueModel
    Set View  = FreeValueForm.Create()

    Result = View.Show(Model)

    Debug.Print "FreeValueForm result: "; Result
    Debug.Print "Model value: "; Model.Value
End Sub