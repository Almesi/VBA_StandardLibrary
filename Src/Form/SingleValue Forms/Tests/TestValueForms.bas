Attribute VB_Name = "TestValueForms"

Option Explicit

Public Sub TestAllValueForms()
    Call TestTreeValueForm
    Call TestColorValueForm
    Call TestCallValueForm
    Call TestEnumValueForm
    Call TestFreeValueForm
End Sub

Public Sub TestTreeValueForm()
    Dim View   As IView
    Dim Model  As ValueModel
    Dim Root   As TreeDisplay
    Dim Group  As TreeDisplay
    Dim Item   As TreeDisplay
    Dim Result As FormReturn

    Set Root = New TreeDisplay
    Root.DisplayName = "Products"
    Root.Value = "products"

    Set Group = New TreeDisplay
    Group.DisplayName = "Fruit"
    Group.Value = "fruit"
    Call Root.AddChild(Group)

    Set Item = New TreeDisplay
    Item.DisplayName = "Apple"
    Item.Value = 101
    Call Group.AddChild(Item)

    Set Item = New TreeDisplay
    Item.DisplayName = "Orange"
    Item.Value = 102
    Call Group.AddChild(Item)

    Set Group = New TreeDisplay
    Group.DisplayName = "Vegetables"
    Group.Value = "vegetables"
    Call Root.AddChild(Group)

    Set Item = New TreeDisplay
    Item.DisplayName = "Carrot"
    Item.Value = 201
    Call Group.AddChild(Item)

    Set Model = New ValueModel
    Set View = TreeValueForm.Create(Root)

    Result = View.Show(Model)

    Debug.Print "TreeValueForm result: "; Result

    If Result = FormReturnOK Then
        Debug.Print "Selected value: "; Model.Value.Value
    Else
        Debug.Print "Selection cancelled."
    End If
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

Public Sub TestColorValueForm()
    Dim View   As IView
    Dim Model  As ValueModel
    Dim Result As FormReturn

    Set Model = New ValueModel
    Model.Value = "FF8000"

    Set View = ColorValueForm.Create()

    Result = View.Show(Model)

    Debug.Print "ColorValueForm result: "; Result
    Debug.Print "Model value: "; Model.Value
End Sub