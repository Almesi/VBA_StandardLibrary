Attribute VB_Name = "TestDataStructures"

Option Explicit

Public Sub TestAllDataStructures()
    Call Run_std_LinkedList_Tests()
    Call Run_std_ParentTree_Tests()
    Call Run_std_Queue_Tests()
    Call Run_std_Stack_Tests()
    Call Run_std_Tree_Tests()
End Sub

Public Sub AssertEqual(ByVal Expected As Variant, ByVal Actual As Variant, ByVal Message As String)
    If Expected <> Actual Then Debug.Print Message 
End Sub

Public Sub AssertTrue(ByVal Condition As Boolean, ByVal Message As String)
    If Not Condition Then Debug.Print Message
End Sub

Public Sub AssertObjectIsNothing(ByVal Value As Object, ByVal Message As String)
    If Not Value Is Nothing Then Debug.Print Message
End Sub