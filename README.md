# VBA_StandardLibrary

This is a Collection of Components designed to ease the work with VBA, by abstraction.
We try to create a library for standard and niche application withing VBA.

## Why use it?
* VBA_StandardLibrary is mainly designed to run without a specific Application inside the VBA enviroment.
* It runs basic functions and many used tools with error handling.
* If you have a wish or request just write a Pull request or open a Discussion.

Currently planned Components can be looked up under Projects

## Short example
```vb
Sub Main()
  
End Sub
```

## Motivation
I first started using VBA at Work to handle basic automation.
Over time my prowess with VBA advanced and i had to work with more complicated topics, those were all different codebases.
Those codebases had some basic tools that i rewrote and specified for the use range like error handling, handling ranges or updating the codebase automatically.
After some time it became tedious to remember all that, which let me to start this library, to standardize all those tools and allow myself to easily manipulate apps out of the VBA enviroment from within VBA.

## What is included
Here a few examples of things included in this library  

Basic programming tools:  

Manipulation of VBE:  

* `std_VBProject.cls`
* 
Excel specific:

* `xl_Workbook.cls`
* `xl_Worksheet.cls`
* 
Files:

* `std_File.cls`
* 
Beyond VBA:

* `std_Graphics` for OpenGL implementation

## Structure
All components will be prefixed by `std_` if they are generic libraries.
Application specific libraries to be prefixed by their name like `xl_` for excel.
