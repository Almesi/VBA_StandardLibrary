Attribute VB_Name = "TestForms"


Option Explicit

Public Sub TestForm()
    Dim ChoicesDataSet      As Table             : Set ChoicesDataSet      = Table.CreateFromTable(ThisWorkbook.Sheets("DB_META_Choices").ListObjects("DB_META_Choices"))
    Dim ChoiceValuesDataSet As Table             : Set ChoiceValuesDataSet = Table.CreateFromTable(ThisWorkbook.Sheets("DB_META_ChoiceValues").ListObjects("DB_META_ChoiceValues"))
    Dim SchemaDataSet       As Table             : Set SchemaDataSet       = Table.CreateFromTable(ThisWorkbook.Sheets("DB_META_Schema").ListObjects("DB_META_Schema"))
    Dim Choices             As ChoiceCollector   : Set Choices             = ChoiceCollector.Create(ChoicesDataSet.DataSet, ChoiceValuesDataSet.DataSet)
    Dim Classifier          As FieldClassifier   : Set Classifier          = FieldClassifier.CreateFromDataSet(SchemaDataSet.DataSet, Choices)
    Dim MyDataBase          As DataBase          : Set MyDataBase          = DataBase.Create(Classifier)
    Dim QueryCommand        As IQueryCommand     : Set QueryCommand        = MakeQueryCommand1()
    Dim MyQuery             As Query             : Set MyQuery             = Query.Create(QueryCommand)

    Call MyQuery.AddSource(NamedDataSet.Create("DB_Transactions", MyDataBase.GetDataSet("DB_Transactions")))
    Call MyQuery.AddSource(NamedDataSet.Create("DB_Categories"  , MyDataBase.GetDataSet("DB_Categories")))
    Call MyQuery.Execute()
    Dim Result            As DataSet         : Set Result            = MyQuery.DataSet
    Call PrintDataSet(Result)

End Sub