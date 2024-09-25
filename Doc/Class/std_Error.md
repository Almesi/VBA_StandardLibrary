# `std_Error.frm`

## Introduction

### What is std_Error?

The main philosophy of std_Error is for the creator of vba projects, who dont want to be bothered with errorhandling.  
Creating a thorough errorhandling system can be tedious. Not having one results in redundant errorhandling of the same type of errors.

It is NOT a tool for programming per se- but more used for end-user input (like the customer)  
Because of that it doesnt stop execution at the line of error, but rather logs it in the desired location  
It is only a tool used to show and handle predefined Errors, those methods are defined in other Components in this Library.


### std_Error consists of 4 main questions
1. What is the Error?
2. Where do i want to show the Error?
3. How do i handle wrong input?
4. What is the Errormessage?



#### 1. Question:
An Error might be wrong input of the programmer, like passing the wrong datatype.  
Another Error might be a too high/low Value of the user.  
3rd Type would be errors, which might result in critical code.  

For these kinds of errors there are several different topics:  
    It needs to be shown to the current user  
    It needs to be logged somewhere for later  
    It needs to be handled by the current user

#### 2. Question:
Everyone has a different taste on this.  
I tried covering all of them:
For Ask/Showing:
1. MsgBox
For Logging:
1. Debug.Print
2. Range.Formula (Excel)
3. File.TextStream.WriteLine
4. Console.PrintEnter (Console.frm)

#### 3. Question:
This is a programmer specific tool
You can manually set the highest value, which is considered a non-critical error (Const Variable `SEVERITY_BREAK`).  
    All errors are considered "maximum error without being severe" unless set otherwise in the Variable "ErrorCatalog"  
One critical error might for example be deleting the whole database  
Either way, once a critical error is detected all code execution will be immediatly dropped  

#### 4. Question:
As i cannot consider every possible Error that might occur and trying so would result in 100.000 or more lines of possible errors `std_Error` is just the implementation of handling.
The actual Functions and Procedures implementing std_Error are the ones doing stuff and handling errors.
For that there is the ErrorCatalog
    If you would like to pass your own Array of Errormessages you can do that.
    You just need a certain setup for Validation: The Catalog needs to be an `Array of (1, n) As Variant` where `Arr(0, n)` is the severity and `Arr(1, n)` is the message
    `Arr(0, 0)` needs to be `ERROR_CATEGORY` and `Arr(1, 0)` needs to be the Class-Name
When trying to implement it in your own Code i would consider doing it like `Sub ProtInit()`, that isnt a must though, as it just saves on runtime.
    

### Example
To show how to use std_Error consider Errorhandling without it:  
THIS IS AN OVER-EXAGGERATION  

```vb
Function CheckUserInput(Value As Variant, MaxValue As Double) As Boolean

    If IsNumeric(Value) = True Then
        If VarType(Value) = vbDouble Then
            If Value = MaxValue Then
                MsgBox("Value1 is equal to Value2" & vbcrlf & "Errortype = System" & vbcrlf & "Errorseverity = x" & vbcrlf & "Value1 = " & Value & vbcrlf & "Value2 = " & MaxValue, vbExlamation, "ERROR")
                Debug.Print "Value1 is equal to Value2" & vbcrlf & "Errortype = System" & vbcrlf & "Errorseverity = x" & vbcrlf & "Value1 = " & Value & vbcrlf & "Value2 = " & MaxValue
            Else
                CheckUserInput = True
            End If
        Else
            MsgBox("Value is not of Type Double" & vbcrlf & "Errortype = System" & vbcrlf & "Errorseverity = x" & vbcrlf & "Value = " & Value, vbExlamation, "ERROR")
            Debug.Print "Value is not of Type Double" & vbcrlf & "Errortype = System" & vbcrlf & "Errorseverity = x" & vbcrlf & "Value = " & Value
        End If
    Else
        MsgBox("Value is not a number" & vbcrlf & "Errortype = System" & vbcrlf & "Errorseverity = x" & vbcrlf & "Value = " & Value, vbExlamation, "ERROR")
        Debug.Print "Value is not of Type Double" & vbcrlf & "Errortype = System" & vbcrlf & "Errorseverity = x" & vbcrlf & "Value = " & Value
    End If
    If ErrorSeverity > ALLOWED_SEVERITY Then
        End
    End If

End Function
```

You would need to this this kind of stuff to every single error.  
With std_Error it would be reduced to:  

```vb
Function CheckUserInput(Value As Variant, MaxValue As Double) As Boolean

    If std_Misc.Number(Value, "=", MaxValue) <> std_Misc.IS_ERROR Then
        CheckUserInput = True
    End If

End Function
```

## Error Functions explained


```vb

' Public Errormethods

    Private       ErrorCatalog(1, 99)   As Variant         ' First  Dimension is Severity/Message. Second Dimension is Index
    Private       Initialized           As Boolean         ' Used once to initialize all Errormessages

    Private Const p_IS_ERROR            As Boolean = True  ' Used to determine, if it is an Error. Dont Change
    Private Const EMPTY_ERROR           As Variant = Empty ' Standard Value if no ErrorValue is passed
    Private Const SEVERITY_BREAK        As Long    = 1000  ' Used to stop the process if this Value is lower than the Error severity. Dont Change
    Private Const ERROR_QUESTION        As Long    = 0001  ' Used to run a Question and handle Error accordingly. Should be below SEVERITY_BREAK. Dont Change
    Private Const ERROR_CATEGORY        As Long    = 0002  ' Needed for ErrorCatalog Validation. Dont Change

    Private       p_ShowError           As Boolean ' Toggle Showing
    Private       p_LogError            As Boolean ' Toggle Logging
    Private       p_LoggingDestination  As Variant ' Pass your goal to log here. Implemented: Excel-Range, Debug.Print, Textfile, Console-Form
    Private       p_ShowDestination     As Variant ' Pass your goal to show here. For example a Range or textfile


    ' Return p_IS_ERROR
    Public Property Get IS_ERROR()

    ' Return p_ShowError
    Public Property Let ShowError(Value As Boolean)

    ' Return p_LogError
    Public Property Let LogError(Value As Boolean)

    ' Return p_LoggingDestination
    Public Property Set LoggingDestination(n_Destination As Variant)

    ' Return p_ShowDestination
    Public Property Set ShowDestination(n_Destination As Variant)


    ' Used to create a new std_Error object with all 4 Set/Let from above
    Public Function Create(Optional n_ShowError As Boolean = True, Optional n_LogError As Boolean = True, Optional n_LoggingDestination As Variant, Optional n_ShowDestination As Variant) As std_Error

    ' Used to Handle Errormessages
    ' Depending on Settings it will Show and Log the Errormessage at the Destination
    ' It will end all Application if the Severity is higher than SEVERITY_BREAK
        ' Catalog is defined as a 2D Array like ErrorCatalog(1, 99). It doesnt have to be this one, it just needs the same Setup. It is used to "import Errormessages from other Components"
        ' Index is the Index of said Catalog
        ' ParamArray ErrorValues are all Values to be displayed aside with the Errormessage 
    Public Function Handle(Catalog() As Variant, Index As Long, ParamArray ErrorValues()) As Boolean

' Private Errormethods
    ' Used to document the errors at the specified destination
        '          Catalog is defined as a 2D Array like ErrorCatalog(1, 99). It doesnt have to be this one, it just needs the same Setup. It is used to "import Errormessages from other Components"
        '          Index is the Index of said Catalog
        ' Optional ErrorValues are all Values to be displayed aside with the Errormessage
    Private Sub Logging(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty)
    
    ' Used to show the errors at the specified destination
        '          Catalog is defined as a 2D Array like ErrorCatalog(1, 99). It doesnt have to be this one, it just needs the same Setup. It is used to "import Errormessages from other Components"
        '          Index is the Index of said Catalog
        ' Optional ErrorValues are all Values to be displayed aside with the Errormessage
    Private Sub Showing(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty)
    
    ' Used to ask the user a yes/no question and will throw an error when NO is pressed, at your specified location
        '          Catalog is defined as a 2D Array like ErrorCatalog(1, 99). It doesnt have to be this one, it just needs the same Setup. It is used to "import Errormessages from other Components"
        '          Index is the Index of said Catalog
        ' Optional ErrorValues are all Values to be displayed aside with the Errormessage
    Private Function Ask(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty) As Boolean

    ' Used to Validate the Catalogs specified in Variable "ErrorCatalog" and will end all execution if false
        ' Catalog is defined as a 2D Array like ErrorCatalog(1, 99). It doesnt have to be this one, it just needs the same Setup. It is used to "import Errormessages from other Components"
    Private Function ValidateCatalog(Catalog() As Variant) As Boolean
    
    ' Used to extract the Message and synthetize all information 
    Private Function GetMessage(Catalog() As Variant, Index As Long, Optional ErrorValues As Variant = Empty) As String

    ' Used to Call ProtInit
    Private Sub Class_Initialize()

    ' Not defined as of now
    Private Sub Class_Terminate()
    
    ' Used to Assign Messages to Me.ErrorCatalog
    Private Sub ProtInit()
'
```
