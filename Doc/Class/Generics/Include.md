# std_Include
### Version 1.0
| X                        | Y                |
| --------                 | -------          |
| Author                   | Almesi           |
| Created                  | 2025-12-04       |
| Last Updated             | 2025-12-04       |
| Related Modules/Classes  | VBGLContext      |
| Tags                     | OOP, VBA, OpenGL |

## Purpose

std_Include manages importing, replacing, removing, and tracking VBA components from external sources.
It abstracts file inclusion, dependency resolution, reference changes, and temporary component lifecycles.


--------------------------------------------------------

## Overview

This class orchestrates dynamic code inclusion into a VBA project. It builds a queue of files, determines the formatter capable of handling each source, and imports or overwrites components as required.
It handles different DataTypes and Input-Sources and Includes them approprietly.
It is able to temporarily include Code for runtime.
A high-level explanation of how it works, key concepts, assumptions, constraints, and known side-effects.
When changing Code in Codemodules you cannot debug through the Code.

**Key capabilities:**
* Importing .bas, .cls, .frm, .txt, and custom formats
* Loading entire folders and optionally recursing into subfolders
* Adding/removing VB project references
* Handling dependencies via multiple IIncludeFormat implementations
* Managing temporary imports that are auto-removed on class termination

**Requirements / Environment:**
* Reference Microsoft Visual Basic for Applications Extensibility 5.3
* Permission for file system operations (FSO)


--------------------------------------------------------

## Properties

| Property    | Type                 | Public | Description |
| --------    | -------              | ------ | -------     |
| Destination |     VBIDE.VBProject  |  True  | Target VBA project where components are included |
| Handler     |     std_ErrorHandler |  True  | Error handling controller |
| Recursive   |     Boolean          |  True  | Whether folder inclusion recurses into subfolders |
| Log         |     Boolean          |  True  | Toggles logging of include/declude operations |
| Overwrite   |     Boolean          |  True  | Overwrites existing component code if found |
| Increment   |     Boolean          |  True  | Enables duplicate inclusion even if component exists |
| Temporary   |     Boolean          |  True  | When enabled, removes components on object destruction |


## Methods
| Methods  | Type         | Public | Description |
| -------- | -------      | ------ | -------     |
| Create                | std_Include | True      | Factory for initializing a new includer with project/handler
| AddFormat             | Void        | True      | Registers a new format handler (strategy pattern)
| Build                 | Boolean     | True      | Executes inclusion queue and processes all items
| AddToQueue            | Boolean     | True      | Places a formatter/source pair into the inclusion queue
| IncludeFolder         | Boolean     | True      | Includes all recognized files within a folder
| IncludeFile           | Boolean     | True      | Adds a file to the inclusion queue
| Declude               | Boolean     | True      | Removes a component from the target project
| OverWriteComponent    | VBComponent | True      | Replaces entire component code
| AddReference          | Boolean     | True      | Adds a VB reference from file or GUID
| RemoveReference       | Boolean     | True      | Removes a VB reference based on attribute type
| IncludeProject        | Boolean     | True      | Adds all components from another VBProject


Further Explanation if nessecary to understand method

--------------------------------------------------------

## Examples:  
1. Describe what the Example is about
```vb
Public Sub Test()
    Dim Shower As IDestination: Set Shower = Nothing
    Dim Logger As IDestination: Set Logger = std_ImmiedeateDestination.Create()
    
    Dim NewErrorHandler As std_ErrorHandler
    Set NewErrorHandler = std_ErrorHandler.Create(Shower, Logger)
    
    Dim Path As String
    Path = "PATH"
    
    Dim Proj As std_Include
    Set Proj = std_Include.Create(ThisWorkbook.VBProject, NewErrorHandler)
    Proj.Recursive = True
    Proj.Log = True
    Proj.Temporary = True
    Proj.Increment = False
    Call Proj.AddFormat(IIncludeFormatVBA)
    Call Proj.AddFormat(IIncludeFormatTXT)
    Call Proj.AddFormat(IIncludeFormatDependencies)

    If Proj.IncludeFolder(Path, "Errorhandling") Then
        If Proj.Build() <> Proj.Handler.IS_ERROR Then
            Debug.Print "Success"
        End If
    End If
End Sub
```

## Extra Information
XXX

## Dependencies
* IIncludeFormat
* std_IncludeSource
* IIncludeFormat-Implementations
* std_ErrorHandler
* std_IncludeFileReader

## Testing
Currently only File Import tested with all available IIncludeFormat Implementations

## Lifecycle Notes
Will declude any temporarily included Component once the object is about to be destroyed

## See Also:
[Name](Path)


# std_IncludeSource
Helper Class for IIncludeFormat.
Is is used as a source holding some information needed for a IIncludeFormat implementation to perform its job.
Might be changed in the Future into another Interface.

# IIncludeFormat
Defines the contract for all format handlers controlling how files or components are included.


| Methods   | Type        | Public | Description |
| --------  | -------     | ------ | -------     |
| Include   | VBComponent | True   | Performs actual inclusion |
| AddQueue  | Boolean     | True   | Adds source to queue using handler-specific logic |
| CanHandle | Boolean     | True   | Determines whether the handler can process given source |
| Code      | String      | True   | Returns file content suitable for overwriting |
| Name      | String      | True   | Returns component name derived from source |


# std_IncludeFormatVBA
Implementation for .bas, .cls and .frm files
# std_IncludeFormatTXT
Implementation for .txt files
# std_IncludeFormatDependencies
Implementation for Dependencies-Files.
Those files just have the name "Dependencies".
Every Line in that File is either an absolute or relative FilePath to another File or Folder that should be implemented.
When Adding this file to a queue it will actually add those other Files and Folders to the queue