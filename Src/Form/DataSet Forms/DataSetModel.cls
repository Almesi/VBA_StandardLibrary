VERSION 1.0 CLASS
BEGIN
    MultiUse = -1
END
Attribute VB_Name = "DataSetModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type TModel
    DataSet         As DataSet
    SelectedRows    As Collection
    SelectedFields  As Collection
End Type

Private This As TModel

Public Property Set DataSet(ByVal       Value  As DataSet)     : Set This.DataSet        = Value               : End Property
Public Property Set SelectedRows(ByVal  Value  As Collection)  : Set This.SelectedRows   = Value               : End Property
Public Property Set SelectedFields(ByVal Value As Collection)  : Set This.SelectedFields = Value               : End Property
Public Property Get DataSet()                  As DataSet      : Set DataSet             = This.DataSet        : End Property
Public Property Get SelectedRows()             As Collection   : Set SelectedRows        = This.SelectedRows   : End Property
Public Property Get SelectedFields()           As Collection   : Set SelectedFields      = This.SelectedFields : End Property
Public Property Get Self()                     As DataSetModel : Set Self                = Me                  : End Property