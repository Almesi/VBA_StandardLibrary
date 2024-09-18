# `std_VBProject.frm`

## Introduction

### What is std_VBProject?

In my Company it is forbidden to use any other programming language but vba
At one point i wasnt able to implement all requierments without the use of .dll files
The problem is, that my companies IT-Security promptly killed my accound when trying to save a workbook with those `Private Declare Function`
My only option was to manually delete the add a component before and delete it after the code is run
This becomes tedious after about 5 minutes, so i started creating this class
The main philosophy of `std_VBProject` is to programmatically handle the Visual Basic Project.
It is inspired by the `#include <iostream>` of C++ 
The only reason for this to be a class is the `Class_Terminate` event.
That is because importing vba code into a project several times will result in duplicates. If the user/programmer does not think about several inclusions it will result in those errors.
To counter that `std_VBProject` allows for temporary inclusion, which will get rid of that code at the end of `Class_Terminate`.
This was the main goal, as this would allow me to use `Private Declare Function` without getting into trouble with IT.



### How to use it
First make sure that Microsoft Visual Basic for Applications Extensibility 5.3 is included into you References
`std_VBProject` allows for several different types of inclusion:
1. Inclusion of Component Files (.bas, .cls, .frm) (exception of .frx files, they cannot be included, but if they are in the same directory as the .frm file it will be included with it)
2. Inclusion of non-Component Files (.txt)
3. Creation of Components and addition of code to it
4. Addition and removal of references
5. Addition of Components from another Project

All of the above can be done temporarely
All of the above can be done with a folder(including all subfolders)

Before you can use any method you first have to use `.Create` and attach a Project



### Extra Information
PasteType has the following modes:
```vb
    PasteType 0 = Include Once
    PasteType 1 = Replace
    PasteType 2 = Include several Times
```
0 will throw an error if the component is already included
1 will try to replace the code in the existing component
2 will include it. If it already exists it will put a number at the end


### Procedures

```vb

Private VBProj As VBIDE.VBProject      ' The Project assignet in .Create or .Project
Private Initialized As Boolean         ' Used to only run initialization once
Private ErrorCatalog(1, 99) As Variant ' 2D Array according to std_Error.cls holding all errormessages
Private TempIncludes() As String       ' 1D Array holding the Name of the Component, which shall be removed at the end of this object´s lifetime


' Public Procedures

    ' Creates a new Object and assignes a VBProject to it
        ' Project = VBProject of desired Worklocation
    Public Function Create(Project As VBIDE.VBProject) As std_VBProject


    ' Used to change the Project at runtime without having to create more than 1 object
        ' Project = VBProject of desired Worklocation
    Public Property Set Project(Project As VBIDE.VBProject)


    ' Used to include a Single File
        '          FilePath      = Whole Name Including FileName and Format
        '          PasteType     = {See Extra Information}
        ' Optional Temporary     = Adds ComponentName to TempIncludes
        ' Optional TryAsInclude  = If True it will it will try to include the code to the component, when the fileformat is not suited for import
        ' Optional ComponentName = Only needed if TryAsInclude is True. Will try to create a new Component, if it doesnt already exist (Watch out for PasteType)
        ' Optional ComponentType = Only needed if TryAsInclude is True. If you know that it should create a new Component you need to tell the function what kind of Component it is. Look up VBA.VBIDE for which numbers you need
    Public Function Include(FilePath As String, Optional PasteType As Long = 0, Optional Temporary As Boolean = False, Optional TryAsInclude As Boolean = False, Optional ComponentName As String = Empty, Optional ComponentType As Long = 0) As Boolean
    

    ' Used to remove the specified Component
        ' CompName = ComponentName to be removed (Doesnt work with the classes needed to run its code)
    Public Function Declude(CompName As String) As Boolean
    

    ' Used to add a new Component to VBProj
        '          Name          = Name of Component
        '          ComponentType = Look up VBA.VBIDE for which numbers you need
        ' Optional PasteType     = {See Extra Information}
        ' Optional ThrowError    = Used to throw an Error
    Public Function AddComponent(Name As String, ComponentType As Long, Optional PasteType As Long = 0, Optional ThrowError As Boolean = True) As Boolean
    

    ' Used to Write all Text from a File to a Specified Component
        '          FilePath  = Whole Name Including FileName and Format
        '          CompName  = Specified Name of Component
        ' Optional StartLine = If Not specified it will Add the Code at the first Line
    Public Function AddCodeFromFile(FilePath As String, CompName As String, Optional StartLine As Long = 1) As Boolean


    ' Used to Write all Text from a Component to another Component
        '          FilePath    = Component-Object from where it will copy the code
        ' Optional ToComponent = Component-Object to where to Paste the Code
        ' Optional StartLine   = If Not specified it will Add the Code at the first Line
    Public Function AddCodeFromComponent(FromComponent As VBComponent, ToComponent As VBComponent, Optional StartLine As Long = 1) As Boolean


    ' Used to Write Text from a String to a Component
        '          CompName  = Name of Component
        '          Text      = String of Chars to be added to Line
        ' Optional StartLine = If Not specified it will Add the Code at the first Line
    Public Function AddCodeFromString(CompName As String, Text As String, Optional StartLine As Long = 1) As Boolean


    ' Used to delete a Part of Code from a Component
        '          CompName  = Name of Component
        ' Optional StartLine = Line where it will start deleting code
        ' Optional EndLine   = Line where it will stop deleting code
    Public Function DeleteCode(CompName As String, Optional StartLine As Long = 1, Optional EndLine As Long = 0) As Boolean
    

    ' Used to Call IncludeFolderArr, but with a ParamArray for Ignore
    Public Function IncludeFolder(FolderPath As String, PasteType As Long, Temporary As Boolean, TryAsInclude As Boolean, ParamArray Ignore() As Variant) As Boolean

    ' Used to include all files from all Subfolders of a folder. Will run .include with the passed parameters
        ' FolderPath    = Whole of Folder without "\" or "/" at the end
        ' PasteType     = {See Extra Information}
        ' Temporary     = Adds ComponentName to TempIncludes
        ' TryAsInclude  = If True it will it will try to include the code to the component, when the fileformat is not suited for import
        ' Ignore        = 1D Array of Text, which represents FilePaths, it will skip over those files
    Public Function IncludeFolderArr(FolderPath As String, PasteType As Long, Temporary As Boolean, TryAsInclude As Boolean, Ignore() As Variant) As Boolean
    

    ' Used to remove all Components, which are found in a folder
    Public Function DecludeFolder(FolderPath As String) As Boolean


    ' Used to add a Reference ( Needs either first Argument OR Arguments 2 to 4)
        ' Optional AddFromFile = FilePath to reference
        ' Optional AddFromGUID = GUID As String (needs argument 3 and 4)
        ' Optional Major       = Major Version of reference (needs argument 2 and 4)
        ' Optional Minor       = Minor Version of reference (needs argument 2 and 3)
    Public Function AddReference(Optional AddFromFile As String = "", Optional AddFromGuid As String = "", Optional Major As Long = 0, Optional Minor As Long = 0) As Boolean


    ' Used to remove all references according to a certain Value
        ' Value       = The Value needed to recognize the reference
        ' AccordingTo = What to search with the Value (eg. Name or GUID)
    Public Function RemoveReference(Value As String, AccordingTo As String) As Boolean


    ' Used to retrieve a Value from a Reference
        '          Ref      = Reference to get the Value from
        ' Optional Property = Which Property to return, default will return everything
    Public Function GetReferenceProperties(Ref As VBIDE.Reference, Optional Property As String = "ALL") As String


    ' Used to get all ComponentNames of a Workbook
        ' Optional WB = Workbook, of which all ComponentNames should be retrieved (if nothing it will use the WB of VBProj)
    Public Function GetAllComponentNames(Optional WB As Workbook = Nothing) As Variant()


    ' Used to Call IncludeFolderArr, but with a ParamArray for Ignore
    Public Function IncludeWorkbook(WB As Workbook, PasteType As Long, Temporary As Boolean, ParamArray Ignore() As Variant) As Boolean
    
    
    ' Used to Include all Components of a Workbook to the Project of VBProj
        ' WB           = Workbook, of which you would like to pass the Project from
        ' PasteType    = {See Extra Information}
        ' Temporary    = Adds ComponentName to TempIncludes
        ' TryAsInclude = If True it will it will try to include the code to the component, when the fileformat is not suited for import
        ' Ignore       = 1D Array of Text, which represents FilePaths, it will skip over those files
    Public Function IncludeWorkbookArr(WB As Workbook, PasteType As Long, Temporary As Boolean, Ignore() As Variant) As Boolean

    ' Used to Include all Components of a Workbook to the Project of VBProj
        '          WB        = Workbook, of which you would like to pass the Project from
        '          CompName  = Name of Component to be included
        '          PasteType = {See Extra Information}
        ' Optional Temporary = Adds ComponentName to TempIncludes
    Public Function IncludeWorkbookSingle(WB As Workbook, CompName As String, PasteType As Long, Optional Temporary As Boolean = False) As Boolean

    ' Used to check if a Value exists and if it should exist
        '          Search      = The Value to be searched for
        ' Optional SearchIn    = String(Object Name) or 1D-Array to search in for the value, if missing then it will use VBProj and search all Components
        ' Optional ShouldExist = XOR´s the ReturnValue
        ' Optional ReturnAsBoolean = If true it will return a Bool, if not it will return an object of the specified search-field (eg. a Component)
    Public Function Exists(Search As Variant, Optional SearchIn As Variant, Optional ShouldExist As Boolean = True, Optional ReturnAsBoolean As Boolean = True) As Variant
'

' Private Procedures
    Private Function Import(FilePath As String, ComponentName As String, Optional PasteType As Long = 0, Optional Temporary As Boolean = False) As Boolean

    ' Used to Loop through all Folders, Subfolders and Files and will include all Files(except those from Ignore)
    Private Function RecursiveInclude(Folder As Object, PasteType As Long, Temporary As Boolean, Ignore As Variant, Optional TryAsInclude As Boolean = False) As Boolean

    ' Tries to return a ComponentName from the specified FilePath 
    Private Function GetCompName(FilePath As String) As String

    ' Checks if FilePath is a .bas, .cls or .frm file
    Private Function CheckFileFormat(FilePath As String) As Boolean

    ' Checks if FilePath is a not importable VBA file (Sheets, .frx, Thisworkbook)
    Private Function CheckClassType(FilePath As String) As Boolean

    ' Simply replaces "\" with "/"
    Private Function ValidateFilePath(FilePath As String) As String

    ' Used to ReDim Preserve TempIncludes and Add a Components Name
    Public Sub AddTemporary(Temporary As Boolean, CompName As String)

    ' Used to Initialize Basic Values and the ErrorCatalog
    Private Sub Class_Initialize()

    ' Used to declude all temporary included Components
    Private Sub Class_Terminate()

    ' Initializes ErrorCatalog
    Private Sub ProtInit()
```


### Examples

Basic Errorhandling
```vb
    Sub Test()

        Dim x As Variant
        Dim y As Variant

        x = 3
        y = 6
        If Console.Number(x, "<", y) = Console.IS_ERROR Then
            MsgBox("HI")
        End If

    End Sub
```

Example for other Console
```vb
    Sub Test2()

        Dim x As Variant
        Dim y As Variant
        PreDec_Answers = Array("y", "n", "m")
        PreDec_Messages = Array("Your Answer is y", "Your Answer is n,", "Your Answer is m")
        Console.Show
        a = CDbl(Console.GetUserInput("Please input a Number for Variable a ", "Double"))
        b = CDbl(Console.GetUserInput("Please input a Number for Variable b ", "Double"))

        If Console.Number(a, "<>", b) = Console.IS_ERROR Then
        End If
        If Console.CheckPredeclaredAnswer("Do you want to add 2 to the variable a?", PreDec_Answers, PreDec_Messages) = "y" Then
            a = a + 2
            Console.PrintEnter "a is: " & a
        End If

    End Sub

    Sub Say(Text As Variant)
        MsgBox(Text)
    End Sub
```
