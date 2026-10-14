VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataSetForm
   Caption         =   "Select Table"
   ClientHeight    =   5000
   ClientWidth     =   7000
   OleObjectBlob   =   "DataSetForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataSetForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Implements IView


'============================================================
'====================[Properties]============================
'============================================================

Private Type TView
    ReturnValue As FormReturn
    Model       As DataSetModel
    DataSet     As DataSet
    AllowEdit   As Boolean
End Type

Private This As TView

'============================================================
'====================[Creation]==============================
'============================================================

Private Sub UserForm_Initialize()
    This.ReturnValue = FormReturnCancel
End Sub

Public Function Create(ByVal nDataSet As DataSet, Optional nAllowEdit As Boolean = True) As DataSetForm
    Set Create = New DataSetForm
    Call Create.Initialize(nDataSet, nAllowEdit)
End Function

Friend Sub Initialize(ByVal nDataSet As DataSet, Optional nAllowEdit As Boolean = True)
    Set This.DataSet = nDataSet
    This.AllowEdit = nAllowEdit
End Sub


'============================================================
'====================[Buttons and Stuff]=====================
'============================================================

Private Sub TableList_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    If This.AllowEdit = False Then Cancel = True
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub


Private Sub OkButton_Click()
    Set This.Model.SelectedRows = New Collection
    Set This.Model.SelectedFields = New Collection

    Dim i As Long
    For i = 1 To TableList.ListItems.Count
        Dim Item As Object
        Set Item = TableList.ListItems.Item(i)
        If Item.Selected Then
            Call This.Model.SelectedRows.Add(i)
        End If
    Next i

    If Not SelectFields() Then
        This.ReturnValue = FormReturnCancel
        Exit Sub
    End If

    This.ReturnValue = FormReturnOK
    Hide
End Sub

Private Function SelectFields() As Boolean
    Dim FieldIndex As Long
    Dim Fields     As Variant
    Dim Value      As Variant
    Dim Text       As String

    Text = Trim(FieldBox.Text)

    If Len(Text) = 0 Then
        For FieldIndex = 1 To This.DataSet.FieldCount
            Call This.Model.SelectedFields.Add(FieldIndex)
        Next FieldIndex

        SelectFields = True
        Exit Function
    End If

    Fields = Split(Text, ",")
    For Each Value In Fields
        Value = Trim(CStr(Value))
        If Len(Value) = 0 Then GoTo NextValue

        If Not IsNumeric(Value) Then
            MsgBox "Invalid field index: " & Value, vbExclamation, "Select Table"
            Exit Function
        End If

        FieldIndex = CLng(Value)

        If FieldIndex < 1 Or FieldIndex > This.DataSet.FieldCount Then
            MsgBox "Field index out of range: " & FieldIndex, vbExclamation, "Select Table"
            Exit Function
        End If

NextValue:
    Next Value

    For Each Value In Fields
        Value = Trim(CStr(Value))
        If Len(Value) > 0 Then
            Call This.Model.SelectedFields.Add(CLng(Value))
        End If
    Next Value
    SelectFields = True
End Function


Private Sub CancelButton_Click()
    OnCancel
End Sub


Private Sub OnCancel()
    This.ReturnValue = FormReturnCancel
    Hide
End Sub

Private Sub Populate()
    Dim RowIndex    As Long
    Dim FieldIndex  As Long
    Dim Entry       As RowEntry
    Dim Field       As FieldEntry
    Dim Item        As Object
    Dim Definition  As FieldDefinition

    TableList.ListItems.Clear
    TableList.ColumnHeaders.Clear

    '--------------------------------------------------
    ' Columns
    '--------------------------------------------------
    Call TableList.ColumnHeaders.Add(, "Row", "Row", 700)
    For FieldIndex = 1 To This.DataSet.FieldCount
        If ShouldShowField(FieldIndex) Then
            Set Definition = This.DataSet.FieldDefinition(FieldIndex)
            Call TableList.ColumnHeaders.Add(, "Field" & FieldIndex, Definition.Name, 1800)
        End If
    Next FieldIndex

    '--------------------------------------------------
    ' Rows
    '--------------------------------------------------
    For RowIndex = 1 To This.DataSet.RowCount
        If ShouldShowRow(RowIndex) Then
            Set Entry = This.DataSet.Row(RowIndex)
            Set Item = TableList.ListItems.Add(, "Row" & RowIndex, CStr(RowIndex))

            Item.Tag = CStr(RowIndex)

            For FieldIndex = 1 To This.DataSet.FieldCount
                If ShouldShowField(FieldIndex) Then
                    If FieldIndex <= Entry.Count Then
                        Set Field = Entry.Item(FieldIndex)

                        Item.SubItems(GetVisibleFieldIndex(FieldIndex)) = CStr(Field.DataEntry.Value)
                    End If
                End If
            Next FieldIndex
        End If
    Next RowIndex
    Call AutoSizeColumns()
End Sub


Private Function ShouldShowRow(ByVal RowIndex As Long) As Boolean
    If This.Model.SelectedRows.Count = 0 Then
        ShouldShowRow = True
        Exit Function
    End If

    ShouldShowRow = CollectionContains(This.Model.SelectedRows, RowIndex)
End Function


Private Function ShouldShowField(ByVal FieldIndex As Long) As Boolean
    If This.Model.SelectedFields.Count = 0 Then
        ShouldShowField = True
        Exit Function
    End If
    ShouldShowField = CollectionContains(This.Model.SelectedFields, FieldIndex)
End Function


Private Function CollectionContains(ByVal Items As Collection, ByVal Value As Long) As Boolean
    Dim Item As Variant

    For Each Item In Items
        If CLng(Item) = Value Then
            CollectionContains = True
            Exit Function
        End If
    Next Item
End Function


Private Function GetVisibleFieldIndex(ByVal FieldIndex As Long) As Long
    Dim Index As Long
    Dim VisibleIndex As Long

    VisibleIndex = 0
    For Index = 1 To This.DataSet.FieldCount
        If ShouldShowField(Index) Then
            VisibleIndex = VisibleIndex + 1
            If Index = FieldIndex Then
                GetVisibleFieldIndex = VisibleIndex
                Exit Function
            End If
        End If
    Next Index
End Function


Private Sub AutoSizeColumns()
    Const CharacterWidth As Long = 5

    Dim ColumnIndex As Long
    Dim RowIndex    As Long
    Dim Width       As Long
    Dim Value       As String

    For ColumnIndex = 1 To TableList.ColumnHeaders.Count
        Width = Len(TableList.ColumnHeaders(ColumnIndex).Text) * CharacterWidth
        For RowIndex = 1 To TableList.ListItems.Count
            If ColumnIndex = 1 Then
                Value = CStr(TableList.ListItems(RowIndex).Text)
            Else
                Value = CStr(TableList.ListItems(RowIndex).SubItems(ColumnIndex - 1))
            End If
            If Len(Value) * CharacterWidth > Width Then
                Width = Len(Value) * CharacterWidth
            End If
        Next RowIndex
        TableList.ColumnHeaders(ColumnIndex).Width = Width
    Next ColumnIndex
End Sub

'============================================================
'====================[Interface]=============================
'============================================================

Public Function IView_Show(ByVal Model As Object) As FormReturn
    Set This.Model = Model
    Populate
    Show
    IView_Show = This.ReturnValue
End Function