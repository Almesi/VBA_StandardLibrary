Attribute VB_Name = "FormBuilderPreBuildTests"

Option Explicit


Public Sub TestAllFormBuilders()
    Dim TestNumber As Long
    Dim ReturnValue As FormReturn

    For TestNumber = 1 To 9
        ReturnValue = TestFormBuilder(TestNumber)
        Debug.Print "Test " & TestNumber & _
                    " returned FormReturn: " & CStr(ReturnValue)
    Next TestNumber
End Sub

Public Function FormLayouter(ByVal Build As IFormBuilder) As IFormBuilder
    Dim Layout As FormLayoutBuilder
    Set Layout = FormLayoutBuilder.Create(Build)
    Layout.AutoSize = True
    
    Dim Result As IFormBuilder
    Set Result = Layout

    Set FormLayouter = Result
End Function

Public Function TestFormBuilder(ByVal TestNumber As Long) As FormReturn
    Dim Form As Object
    Set Form = New EmptyForm

    Dim Builder As IFormBuilder

    Select Case TestNumber
        Case 1: Set Builder = TestGenericInput(Form)
        Case 2: Set Builder = TestYesNoQuestion()
        Case 3: Set Builder = TestTenFieldForm()
        Case 4: Set Builder = TestRepeatedTextBoxes()
        Case 5: Set Builder = TestMixedControlRow()
        Case 6: Set Builder = TestDateRangeField()
        Case 7: Set Builder = TestAddressField()
        Case 8: Set Builder = TestLabeledActions()
        Case 9: Set Builder = TestLoginForm()
        Case Else
            Err.Raise vbObjectError + 1000, "TestFormBuilder", _
                      "Unknown test number: " & TestNumber
    End Select

    Dim Layouter As IFormBuilder
    Set Layouter = FormLayouter(Builder)

    Dim Result As FormBuilderContext
    Set Result = FormBuilderContext.Create(FormBuilderForm.Create(Form))

    Call Layouter.Build(Result, Result.Scope)

    Dim View As IView
    Set View = Form
    TestFormBuilder = View.Show(Nothing)
End Function


'=========================================================================
' 1. GENERIC INPUT
'
'   Value
'   [ ValueTextBox ]
'
'   [ OKButton ] [ CancelButton ]
'=========================================================================

Public Function TestGenericInput(ByVal Form As Object) As IFormBuilder
    Dim View            As IView        : Set View           = Form
    Dim TextBoxFactory  As Object       : Set TextBoxFactory = CreateTextBoxFactory("Value")
    Dim ValueBuilder    As IFormBuilder : Set ValueBuilder   = LabelWithColumn(FormBuilderControl(TextBoxFactory))
    Dim OKFactory       As Object       : Set OKFactory      = CreateButtonFactory("OKButton", "OK")
    Dim CancelFactory   As Object       : Set CancelFactory  = CreateButtonFactory("CancelButton", "Cancel")    
    Dim ButtonsBuilder  As IFormBuilder : Set ButtonsBuilder = OKCancelButtons(OKFactory, CancelFactory)
            
    Set OKFactory.Functions     = MakeFunctionDirSingle("Click", BoundFunction.Create(NamedFunction.Create("OKButtonTest1", False, False), View))
    Set CancelFactory.Functions = MakeFunctionDirSingle("Click", BoundFunction.Create(NamedFunction.Create("CancelButtonTest1", False, False), View))

    Set TestGenericInput = _
        FormBuilderColumn( _
            FormBuilderSequence( _
                ValueBuilder, _
                ButtonsBuilder _
            ), _
            0, _
            0 _
        )

End Function


'=========================================================================
' 2. YES / NO QUESTION WITH LABEL
'
'   Do you want to continue?
'
'   [ Yes ] [ No ]
'=========================================================================

Public Function TestYesNoQuestion() As IFormBuilder
    Dim QuestionLabelFactory As Object       : Set QuestionLabelFactory = CreateLabelFactory("ContinueQuestion", "Do you want to continue?")
    Dim QuestionBuilder      As IFormBuilder : Set QuestionBuilder      = FormBuilderControl(QuestionLabelFactory)
    Dim YesFactory           As Object       : Set YesFactory           = CreateButtonFactory("Yes", "Yes")
    Dim NoFactory            As Object       : Set NoFactory            = CreateButtonFactory("No", "No")

    Dim ButtonsBuilder As IFormBuilder
    Set ButtonsBuilder = YesNoButtons( _
        YesFactory, _
        NoFactory _
    )

    Set TestYesNoQuestion = _
        FormBuilderColumn( _
            FormBuilderSequence( _
                QuestionBuilder, _
                ButtonsBuilder _
            ), _
            0, _
            10 _
        )

End Function

'=========================================================================
' 3. TEN-FIELD FORM
'
'   First Name            Last Name            Email             etc.
'   [ FirstNameTextBox ]  [ LastNameTextBox ]  [ EmailTextBox ]  etc.

'=========================================================================

Public Function TestTenFieldForm() As IFormBuilder
    Dim FirstNameBuilder As IFormBuilder : Set FirstNameBuilder = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("FirstName")) , CreateLabelFactory("FirstName" , "First Name"))
    Dim LastNameBuilder  As IFormBuilder : Set LastNameBuilder  = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("LastName"))  , CreateLabelFactory("LastName"  , "Last Name"))
    Dim EmailBuilder     As IFormBuilder : Set EmailBuilder     = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("Email"))     , CreateLabelFactory("Email"     , "Email"))
    Dim PhoneBuilder     As IFormBuilder : Set PhoneBuilder     = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("Phone"))     , CreateLabelFactory("Phone"     , "Phone"))
    Dim AddressBuilder   As IFormBuilder : Set AddressBuilder   = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("Address"))   , CreateLabelFactory("Address"   , "Address"))
    Dim CityBuilder      As IFormBuilder : Set CityBuilder      = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("City"))      , CreateLabelFactory("City"      , "City"))
    Dim StateBuilder     As IFormBuilder : Set StateBuilder     = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("State"))     , CreateLabelFactory("State"     , "State"))
    Dim ZIPBuilder       As IFormBuilder : Set ZIPBuilder       = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("ZIP"))       , CreateLabelFactory("ZIP"       , "ZIP"))
    Dim CountryBuilder   As IFormBuilder : Set CountryBuilder   = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("Country"))   , CreateLabelFactory("Country"   , "Country"))
    Dim CompanyBuilder   As IFormBuilder : Set CompanyBuilder   = LabelWithColumn(FormBuilderControl(CreateTextBoxFactory("Company"))   , CreateLabelFactory("Company"   , "Company"))

    Set TestTenFieldForm = _
        FormBuilderRow( _
            FormBuilderSequence( _
                FirstNameBuilder, _
                LastNameBuilder, _
                EmailBuilder, _
                PhoneBuilder, _
                AddressBuilder, _
                CityBuilder, _
                StateBuilder, _
                ZIPBuilder, _
                CountryBuilder, _
                CompanyBuilder _
            ), _
            0, _
            8 _
        )

End Function

'=========================================================================
' 4. THREE IDENTICAL CONTROLS IN A ROW
'
'   [ TextBox ] [ TextBox ] [ TextBox ]
'=========================================================================

Public Function TestRepeatedTextBoxes() As IFormBuilder
    Dim TextBoxFactory As Object
    Set TextBoxFactory = CreateTextBoxFactory("Input")

    Dim TextBoxesBuilder As IFormBuilder
    Set TextBoxesBuilder = ControlRow( _
        TextBoxFactory, _
        3, _
        5 _
    )

    Set TestRepeatedTextBoxes = TextBoxesBuilder

End Function

'=========================================================================
' 5. THREE DIFFERENT CONTROLS IN A ROW
'
'   [ TextBox ] [ TextBox ] [ Button ]
'=========================================================================

Public Function TestMixedControlRow() As IFormBuilder
    Dim FirstFactory       As Object : Set FirstFactory  = CreateTextBoxFactory("First")
    Dim SecondFactory      As Object : Set SecondFactory = CreateTextBoxFactory("Second")
    Dim ActionFactory      As Object : Set ActionFactory = CreateButtonFactory("Action", "Action")

    Dim RowBuilder As IFormBuilder
    Set RowBuilder = ControlsRow( _
        FirstFactory, _
        SecondFactory, _
        ActionFactory _
    )

    Set TestMixedControlRow = RowBuilder

End Function

'=========================================================================
' 6. FIELD WITH MULTIPLE CONTROLS
'
'   Date Range
'   [ From ] [ To ]
'=========================================================================

Public Function TestDateRangeField() As IFormBuilder
    Dim FromFactory As Object : Set FromFactory = CreateTextBoxFactory("FromDate")
    Dim ToFactory   As Object : Set ToFactory   = CreateTextBoxFactory("ToDate")

    Dim DateRangeBuilder As IFormBuilder
    Set DateRangeBuilder = LabelWithRow( _
        FormBuilderControls( _
            FromFactory, _
            ToFactory _
        ), _
        CreateLabelFactory( _
            "DateRange", _
            "Date Range" _
        ), _
        5 _
    )

    Set TestDateRangeField = DateRangeBuilder

End Function

'=========================================================================
' 7. FIELD WITH MULTIPLE CONTROLS VERTICALLY
'
'   Address
'   [ Street ]
'   [ City   ]
'   [ ZIP    ]
'=========================================================================

Public Function TestAddressField() As IFormBuilder
    Dim StreetFactory      As Object : Set StreetFactory      = CreateTextBoxFactory("Street")
    Dim CityFactory        As Object : Set CityFactory        = CreateTextBoxFactory("City")
    Dim ZIPFactory         As Object : Set ZIPFactory         = CreateTextBoxFactory("ZIP")

    Dim ContentBuilder As IFormBuilder
    Set ContentBuilder = FormBuilderColumn( _
        FormBuilderControls( _
            StreetFactory, _
            CityFactory, _
            ZIPFactory _
        ), _
        0, _
        5 _
    )

    Set TestAddressField = LabelWithColumn( _
        ContentBuilder, _
        CreateLabelFactory( _
            "Address", _
            "Address" _
        ), _
        8 _
    )

End Function

'=========================================================================
' 8. LABELED BUTTON GROUP
'
'   Actions    [ Save ] [ Cancel ] [ Reset ]
'=========================================================================

Public Function TestLabeledActions() As IFormBuilder
    Dim SaveFactory        As Object : Set SaveFactory        = CreateButtonFactory("Save", "Save")
    Dim CancelFactory      As Object : Set CancelFactory      = CreateButtonFactory("Cancel", "Cancel")
    Dim ResetFactory       As Object : Set ResetFactory       = CreateButtonFactory("Reset", "Reset")

    Dim ActionsBuilder As IFormBuilder
    Set ActionsBuilder = LabeledButtonGroup( _
        CreateLabelFactory( _
            "Actions", _
            "Actions" _
        ), _
        SaveFactory, _
        CancelFactory, _
        ResetFactory _
    )

    Set TestLabeledActions = ActionsBuilder

End Function

'=========================================================================
' 9. COMPLETE LOGIN FORM
'
'   Username
'   [ UsernameTextBox ]
'
'   Password
'   [ PasswordTextBox ]
'
'   [ Login ] [ Cancel ]
'=========================================================================

Public Function TestLoginForm() As IFormBuilder
    Dim UsernameFactory    As Object
    Set UsernameFactory    = CreateTextBoxFactory("Username")
    Dim UsernameBuilder As IFormBuilder
    Set UsernameBuilder = LabelWithColumn( _
        FormBuilderControl(UsernameFactory), _
        CreateLabelFactory( _
            "Username", _
            "Username" _
        ) _
    )

    Dim PasswordFactory    As Object
    Set PasswordFactory    = CreateTextBoxFactory("Password")
    Dim PasswordBuilder As IFormBuilder
    Set PasswordBuilder = LabelWithColumn( _
        FormBuilderControl(PasswordFactory), _
        CreateLabelFactory( _
            "Password", _
            "Password" _
        ) _
    )

    Dim LoginFactory       As Object : Set LoginFactory       = CreateButtonFactory("Login", "Login")

    Dim CancelFactory      As Object : Set CancelFactory      = CreateButtonFactory("Cancel", "Cancel")

    Dim ButtonsBuilder As IFormBuilder
    Set ButtonsBuilder = OKCancelButtons( _
        LoginFactory, _
        CancelFactory _
    )

    Set TestLoginForm = _
        FormBuilderColumn( _
            FormBuilderSequence( _
                UsernameBuilder, _
                PasswordBuilder, _
                ButtonsBuilder _
            ), _
            0, _
            10 _
        )

End Function