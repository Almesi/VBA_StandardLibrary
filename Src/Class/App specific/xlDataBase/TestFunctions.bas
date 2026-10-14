Attribute VB_Name = "TestFunctions"


Option Explicit

Public Function CreateDataBase(ByVal WB As Workbook) As DataBase
    Dim ChoicesDataSet      As Table             : Set ChoicesDataSet      = Table.CreateFromTable(WB.Sheets("DB_META_Choices").ListObjects("DB_META_Choices"))
    Dim ChoiceValuesDataSet As Table             : Set ChoiceValuesDataSet = Table.CreateFromTable(WB.Sheets("DB_META_ChoiceValues").ListObjects("DB_META_ChoiceValues"))
    Dim SchemaDataSet       As Table             : Set SchemaDataSet       = Table.CreateFromTable(WB.Sheets("DB_META_Schema").ListObjects("DB_META_Schema"))
    Dim Choices             As ChoiceCollector   : Set Choices             = ChoiceCollector.Create(ChoicesDataSet.DataSet, ChoiceValuesDataSet.DataSet)
    Dim Classifier          As FieldClassifier   : Set Classifier          = FieldClassifier.CreateFromDataSet(SchemaDataSet.DataSet, Choices)
    Dim MyDataBase          As DataBase          : Set MyDataBase          = DataBase.Create(Classifier)

    Set CreateDataBase = MyDataBase
End Function

Public Sub TestFunction()
    Dim MyDataBase          As DataBase          : Set MyDataBase          = CreateDataBase(ThisWorkbook)
    Dim QueryCommand        As IQueryCommand     : Set QueryCommand        = MakeQueryCommand1()
    Dim MyQuery             As Query             : Set MyQuery             = Query.Create(QueryCommand)

    Call MyQuery.AddSource(NamedDataSet.Create("DB_Transactions", MyDataBase.GetDataSet("DB_Transactions")))
    Call MyQuery.AddSource(NamedDataSet.Create("DB_Categories"  , MyDataBase.GetDataSet("DB_Categories")))
    Call MyQuery.Execute()
    Dim Result            As DataSet         : Set Result            = MyQuery.DataSet
    Call PrintDataSet(Result)

End Sub

Public Sub PrintDataSet(ByVal DataSet As DataSet)
    On Error GoTo ErrorHandler

    Dim Field       As Field
    Dim Entry       As DataEntry
    Dim i           As Long
    Dim MaxRows     As Long
    Dim ColWidths() As Long
    Dim Values()    As String

    ' Find maximum number of entries
    MaxRows = 0

    For Each Field In DataSet.Fields
        If Field.DataEntries.Count > MaxRows Then
            MaxRows = Field.DataEntries.Count
        End If
    Next Field

    ' Determine number of fields
    Dim FieldCount As Long
    FieldCount = DataSet.Fields.Count

    ReDim ColWidths(FieldCount - 1)
    ReDim Values(FieldCount - 1, MaxRows)

    ' Collect values and determine column widths
    Dim Col As Long
    Dim Row As Long

    Col = 0

    For Each Field In DataSet.Fields
        Values(Col, 0) = Field.Definition.Name
        If Len(Values(Col, 0)) > ColWidths(Col) Then
            ColWidths(Col) = Len(Values(Col, 0))
        End If
        Row = 1
        For Each Entry In Field.DataEntries
            Values(Col, Row) = CStr(Entry.Value)
            If Len(Values(Col, Row)) > ColWidths(Col) Then
                ColWidths(Col) = Len(Values(Col, Row))
            End If
            Row = Row + 1
        Next Entry
        Col = Col + 1
    Next Field

    ' Print row by row
    For Row = 0 To MaxRows
        For Col = 0 To FieldCount - 1
            If Col > 0 Then
                Debug.Print "  ";
            End If
            Debug.Print PadRight(Values(Col, Row), ColWidths(Col));
        Next Col
        Debug.Print
    Next Row
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR: " & Err.Number & " - " & Err.Description
End Sub

Private Function PadRight(ByVal Text As String, ByVal Width As Long) As String
    PadRight = Text & Space$(Width - Len(Text))
End Function

Public Function MakeQueryCommand1() As IQueryCommand
    Dim Source   As QuerySourceCommand : Set Source   = QuerySourceCommand.Create("DB_Transactions")
    Dim Selector As QuerySelectCommand : Set Selector = QuerySelectCommand.Create("ID", "Description", "DB_Categories::ID", "Amount", "Transaction Type")
    Dim Filter   As QueryWhereCommand  : Set Filter   = QueryWhereCommand.Create("DB_Categories::ID", "=", 15)
    Dim Adding   As QueryAddCommand    : Set Adding   = QueryAddCommand.Create("Filtered Food")
    Dim Union    As QueryUnionCommand  : Set Union    = QueryUnionCommand.Create("Filtered Food")
    Dim Start    As QueryCommandChain  : Set Start    = QueryCommandChain.Create(Source, Selector, Filter, Adding, Union)
    Set MakeQueryCommand1 = Start
End Function