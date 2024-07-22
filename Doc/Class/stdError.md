# `stdError.frm`

## Introduction

### What is stdError?

The main philosophy of stdError is for the creator of vba projects, who dont want to be bothered with errorhandling.
Creating a thorough errorhandling system can be tedious. Not having one results in redundant errorhandling of the same type of errors.

It is NOT a tool for programming per se- but more used for end-user input (like the customer)
Because of that it doesnt stop execution at the line of error, but rather logs it in the desired location



### stdError consists of 3 main questions

    1. What is the Error?
    2. Where do i want to show the Error?
    3. How do i handle wrong input?



#### 1. Question:
    An Error might be wrong input of the programmer, like passing the wrong datatype.
    Another Error might be a too high/low Value of the user.
    3rd Type would be errors, which might result in critical code.

    For these kinds of errors there are several different topics:
        It needs to be shown to the current user
        It needs to be logged somewhere for later
        It needs to be fixed by the current user

#### 2. Question:
    There are different tastes to this.
    I tried covering all of them:
    One Method is the "standard" one MsgBox and Debug.Print. No further loss of time, because they are directly tied and well established in VBA.
        Problem is logging. You cant look into it once the code is run
    Another Method i would call the "c++" method, where programmers prefer the command prompt
        For that reason stdError includes the console, a tool for run-time manipulation. Further down the line i will explain it further
        Problem is the same as the previous method
    Third method is printing it to a file/directly into a ms-tool like word/excel/outlook
        Through this way you can look into the errors even after everything is run and closed

#### 3. Question:
    This is a programmer specific tool
    You can manually set the highest value, which is considered a non-critical error (Const Variable "SEVERITY_BREAK").
        All errors are considered "maximum error without being severe" unless set otherwise in the Variable "ErrorCatalog"
    One critical error might for example be deleting the whole database
    Either way, once a critical error is detected all code execution will be immediatly dropped
    

### Example
    To show how to use stdError consider Errorhandling without it:
    THIS IS AN OVER-EXAGGGERATION

```vb
Function CheckUserInput(Value As Variant, MaxValue As Double) As Boolean

    If IsNumeric(Value) = True Then
        If VarType(Value) = vbDouble Then
            If Value = MaxValue Then
                MsgBox("Value1 is equal to Value2" & Chr(13) & Chr(10) & "Errortype = System" & Chr(13) & Chr(10) & "Errorseverity = x" & Chr(13) & Chr(10) & "Value1 = " & Value & Chr(13) & Chr(10) & "Value2 = " & MaxValue, vbExlamation, "ERROR")
                Debug.Print "Value1 is equal to Value2" & Chr(13) & Chr(10) & "Errortype = System" & Chr(13) & Chr(10) & "Errorseverity = x" & Chr(13) & Chr(10) & "Value1 = " & Value & Chr(13) & Chr(10) & "Value2 = " & MaxValue
            Else
                CheckUserInput = True
            End If
        Else
            MsgBox("Value is not of Type Double" & Chr(13) & Chr(10) & "Errortype = System" & Chr(13) & Chr(10) & "Errorseverity = x" & Chr(13) & Chr(10) & "Value = " & Value, vbExlamation, "ERROR")
            Debug.Print "Value is not of Type Double" & Chr(13) & Chr(10) & "Errortype = System" & Chr(13) & Chr(10) & "Errorseverity = x" & Chr(13) & Chr(10) & "Value = " & Value
        End If
    Else
        MsgBox("Value is not a number" & Chr(13) & Chr(10) & "Errortype = System" & Chr(13) & Chr(10) & "Errorseverity = x" & Chr(13) & Chr(10) & "Value = " & Value, vbExlamation, "ERROR")
        Debug.Print "Value is not of Type Double" & Chr(13) & Chr(10) & "Errortype = System" & Chr(13) & Chr(10) & "Errorseverity = x" & Chr(13) & Chr(10) & "Value = " & Value
    End If
    If ErrorSeverity > ALLOWED_SEVERITY Then
        End
    End If

End Function
```

    You would need to this this kind of stuff to every single error.
    With stdError it would be reduced to:

```vb
Function CheckUserInput(Value As Variant, MaxValue As Double) As Boolean

    If Console.Number(Value, "=", MaxValue) = Console.IS_ERROR Then
        CheckUserInput = True
    End If

End Function
```



## Console
Another big part is the console
Its used to log errors, print success or failure, ask the user for input, decide process with yes/no questions and run procedures (macros) and extras


### Preparation
To Prepare the console you have to do the following:
Go to stdError and search for the function `LogMode`, there you have to write the number of corresponding to `LogModeEnum`
Run `Console.Show`
    This will initialize the console

Now the Console can be used in process

#### 1. Log Errors
When you use any Errorhandling function of stdError and the Console is activated then the error will be printed to the console

#### 2. User Input
There are 2 main User-interactions-
    One is a message, followed by predeclared answers like yes/no,maybe or anything the programmer would like to use
        VBA will continue to run until you write any of the available answers, else it will print, that your value is not allowed
        If the input is allowed a user defined message may be shown
    The second one is a message, where it will ask you for a value of the user.
        This might be combined with further errorhandling like checking if the input is of the right datatype, a wrong input will be shown accordingly

#### 3. Running macros
This tool is strictly defined:
    Write a Variable name
    Enter |; |, as the seperator of arguments
    Up to 29 additional arguments are allowed (1st argument is ProcedureName, all following arguments are arguments for said procedure)

#### 4. Extras
As of now there are the following special commands
    Help
        This will print a text explaining the functionalities of the console

### Extra Information
    The console works with special modes:
```vb
     Private Enum WorkModeEnum
        Logging = 0
        UserInputt = 1
        PreDeclaredAnswer = 2
        UserLog = 3
    End Enum
```
Logging is the basic one, where the console only recieves information
UserInputt is variable input of the user
PreDeclaredAnswer is for predeclared answers (duh)
UserLog is for running procedures and extras


## Console Functions explained

```vb

' Public Console Functions

    ' Shows defined Message on Console
    ' Runs in loop
    ' Checks for userinput
    ' once typed check ifs of defined datatype
    ' keeps running until an acceptable input is inserted
    Public Function GetUserInput(Message As Variant, Optional InputType As String = "VARIANT") As Variant


    ' Shows defined Message on Console
    ' Runs in loop
    ' Checks for userinput
    ' once typed check ifs of allowed value
    ' keeps running until an acceptable input is inserted
    ' print successmessage (optional)
    Public Function CheckPredeclaredAnswer(Message As Variant, AllowedValues As Variant, Optional Answers As Variant = Empty) As Variant


    ' Returns a String of the workbook-folderpath with a "Recognizer"
    Public Function PrintStarter() As Variant

    ' Prints a message to the console and inserts a new line
    Public Sub PrintEnter(Text As Variant)

    'Prints a message to the console without inserting a new line
    Public Sub PrintConsole(Text As Variant)



' Private Console Functions


    ' Initializes the first text and inserts the Printstarter
    Private Sub UserForm_Initialize()

    ' handles userinput buttons
        'Enter key will run HandleEnter
        'Up key will insert the previous line
        'Down key will insert the next line
    Private Sub ConsoleText_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    ' Gets a specified line and removes everything before the Recognizer
    Private Function GetLine(Text As String, Index As Long) As String

    ' When running enter the system will check what it should do
        ' Running in Userinputt mode will get the userinput and removes the message, which is usually printed in the same line
        ' Running in UserLog will remove everything before the Recognizer
            ' Then it will split the text into arguments for HandleCode
        ' Running in Logging will insert a new line and print the error
    Private Sub HandleEnter()

    ' If the system recognizes the line as a procedure it should run, this function runs
    ' If it recognizes the line as a special, then it will run HandleSpecial
    ' Runs the specified code and if all goes right it will print "success" after running the function, else it will show an error message
    Private Function HandleCode(ParamArray Arguments() As Variant) As Variant

    ' Looks, if the first argument passed is of a special type
    ' These extras are explained further up this file
    Private Function HandleSpecial(ParamArray Arguments() As Variant)

    ' Gets a string of the message which will be displayed when initializing the Console
    Private Function GetStartText() As String

    ' Prints a string to the console, explaining the functionality of the console
    Private Sub HandleHelp()

```




## Error Functions explained


```vb

' Public Errormethods

    ' Returns the p_IS_ERROR Constant
    ' This exists for better readability for the user, by not having to worry about "Is true the error or is false the error?"
    Public Property Get IS_ERROR()
        
    ' Defines if Errors should be shown at the desired location to the user or not
    Public Property Let ShowError(Value As Boolean)

    ' Defines if Errors should be printed to the desired locationor not
    Public Property Let PrintError(Value As Boolean)


    ' Handles the error
    ' This includes:
    ' Checking if the error should be shown, printed or the user should be asked if he wants to continue the process
    ' Ends the execution if Severity is too high
    ' Errorcategory is the type of error
    ' Errorindex is the row, where the message is held
    ' The following errorvalues are optional and will be displayed accordingly
    Public Function Handle(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR) As Boolean

    ' Can check the following:
        ' Is Firstvalue Empty?
        ' Is Firstvalue like SecondValue according to defined operator?
        ' Is FirstValue smaller than MinValue?   (Can be achieved with the first 3 arguments, but to include it in one statement this argument exists)
        ' Is FirstValue bigger than MaxValue?    (Can be achieved with the first 3 arguments, but to include it in one statement this argument exists)
    Public Function Variable(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean

    ' Can check the following:
        ' Is Firstvalue Nothing?
        ' Firstvalue like SecondValue according to defined operator?
    Public Function Object(FirstValue As Object, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Object = EMPTY_ERROR) As Boolean

    ' Checks if Workbook is open and if it should do that or not
    Public Function Workbook(WorkbookName As String, Optional ShouldExist As Boolean = True) As Boolean

    ' Checks if Worksheet is open and if it should do that or not (runs Workbook)
    Public Function Worksheet(WorkbookName As String, SheetName As String, Optional ShouldExist As Boolean = True) As Boolean

    ' Compare Strings with a passed Operator
    Public Function Strings(Text As String, Operator As String, SecondText As String) As Boolean

    ' Superset of Variable, used to check if it is a number and Handle Errors
    Public Function Number(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean

    ' Superset of Variable, used to check if it is a date and Handle Errors
    Public Function Dates(FirstValue As Variant, Optional Operator As String = EMPTY_ERROR, Optional SecondValue As Variant = EMPTY_ERROR, Optional MinValue As Variant = EMPTY_ERROR, Optional MaxValue As Variant = EMPTY_ERROR) As Boolean

    ' Checks if File exists and if it should do that or not
    Public Function File(FilePath As String, Optional ShouldExist As Boolean = True) As Boolean

    ' Check Connection to specified Computer and Handle Errors
    Public Function Connection(Optional Computer As String = ".", Optional ShouldExist As Boolean = True) As Boolean

    ' Check DatabaseConnection and Handle Errors
    Public Function ConnectToDatabase(DataBasePath As String) As Boolean

    ' Check if passed Variable is InputType according to ShouldBe
    Public Function DataType(Value As Variant, InputType As String, Optional ShouldBe As Boolean = True) As Boolean
'



' Private Errormethods
    ' Print Error to designated Location (Log)
    Private Sub Printt(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR)
    
    ' Show Error to designated Location
    Private Sub Showw(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR)
    
    ' Asks Yes/No Question (No will raise an Error)
    Private Function Ask(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR) As Boolean
    
    ' Gets Errormessage
    Private Function GetMessage(ErrorCategory As Long, ErrorIndex As Long, Optional ErrorValue1 As Variant = EMPTY_ERROR, Optional ErrorValue2 As Variant = EMPTY_ERROR, Optional ErrorValue3 As Variant = EMPTY_ERROR, Optional ErrorValue4 As Variant = EMPTY_ERROR) As String

    ' Gets Name of Errorcategory
    Private Function GetCategory(ErrorCategory As Long) As String

    ' Runs once to Initialize all Errormessages
    Private Sub ProtInit()
'
```