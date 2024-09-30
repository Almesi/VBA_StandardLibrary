Hello, and thanks for looking into contributing to this library!

In this file I want to document the main ideas behind the library.


## Other Contributions

If possible try to contribute to sancarn´s stdVBA first, as he is trying to create a standard library for vba.


## Installation

The components can be dragged from windows explorer and dropped into the VBA project.
The component `std_Error.cls` is a requierement for most components.


## Philosophy

This library does NOT look for a specific type of component. Modules, Classes and Forms are all welcome.
It is meant as programming tools completely executed withing vba or an api between vba-using apps.
This library is meant for save programming, meaning all functions should consider all possible errors and handle. them, without throwing those error(debug window), for that `std_Error.cls` is the dependency for most components, to handle all the errors.
This library is "simple datatypes first", meaning, it tries to use string, doubles, integers, booleans etc.
If possible try not to pass objects.
Main reason for that is, that the object should be described/called via its name, a key or index, so that you can call it via a written text e.g. a console.

## Locality
Please try to keep your Public and Global Variables and Function to a minimum, especially when it is a module and not a class/form


## Class/Module naming

To not confuse these component with the ones from sancarn these are labeled as `std_XXX` instead of `stdXXX`.
If it is only used in certain apps please use a shortform of the app(s) for example `xl_XXX` for Excel Or `xlpp_XXX` for Excel and Powerpoint.


## Multi-application

This library should try to work for every vba-using app.
Modules should avoid using features which are only available for certain apps, except when they only work in one/are designed to work in certain apps(For example an Excel-Access API).


## Dependencies and ease of use

If possible try to keep the dependencies as low as possible.
For that please use PredeclaredID = True for classes and forms.
Reason is, that the component should be imported and work on the fly.
In that philosophy also please try to keep all the code for a set of functions in 1 component(less dependencies).


## Documentation

The most important bit is the documentation.
Im trying to document everything in a dedicated `.md` file.
Documentation in Code is not needed, but that would be no problem, im not doing it for better readability.
Documentation need to at least have a description of what and how it does it and why it is needed.
All arguments also have to be desribed why they are there, what they are used for and extras that happen at certain values.

For example:

````vb
    ' Used to include a Single File. Does this by checking if it can be imported of added via codemodule
        '          FilePath      = Whole Name Including FileName and Format
        '          PasteType     = {See Extra Information}
        ' Optional Temporary     = Adds ComponentName to TempIncludes
        ' Optional TryAsInclude  = If True it will it will try to include the code to the component, when the fileformat is not suited for import
        ' Optional ComponentName = Only needed if TryAsInclude is True. Will try to create a new Component, if it doesnt already exist (Watch out for PasteType)
        ' Optional ComponentType = Only needed if TryAsInclude is True. If you know that it should create a new Component you need to tell the function what kind of Component it is. Look up VBA.VBIDE for which numbers you need
    Public Function Include(FilePath As String, Optional PasteType As Long = 0, Optional Temporary As Boolean = False, Optional TryAsInclude As Boolean = False, Optional ComponentName As String = Empty, Optional ComponentType As Long = 0, Optional ThrowError As Boolean = True) As Boolean
````

As many people say: the code is the documentation.
For that please try to write out your Variable names. Of course, if the code gets too long then try to shorten it or use short names.
As long as it is easily readable and understandable it should be fine.