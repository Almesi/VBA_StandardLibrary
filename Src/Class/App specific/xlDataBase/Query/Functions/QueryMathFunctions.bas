Attribute VB_Name = "QueryMathFunctions"

Option Explicit

Public Function QuerySum(ByVal Values As Collection) As Double
    Dim Entry As DataEntry
    Dim Result As Double

    For Each Entry In Values
        Result = Result + CDbl(Entry.Value)
    Next Entry

    QuerySum = Result
End Function

Public Function QueryAverage(ByVal Values As Collection) As Double
    Dim Entry As DataEntry
    Dim Total As Double
    Dim Count As Long

    For Each Entry In Values
        Total = Total + CDbl(Entry.Value)
        Count = Count + 1
    Next Entry

    QueryAverage = Total / Count
End Function

Public Function QueryCount(ByVal Values As Collection) As Long
    QueryCount = Values.Count
End Function

Public Function QueryMin(ByVal Values As Collection) As Double
    Dim Entry As DataEntry
    Dim Result As Double
    Dim IsFirst As Boolean

    For Each Entry In Values
        If Not IsFirst Then
            Result = CDbl(Entry.Value)
            IsFirst = True
        ElseIf CDbl(Entry.Value) < Result Then
            Result = CDbl(Entry.Value)
        End If
    Next Entry

    QueryMin = Result
End Function

Public Function QueryMax(ByVal Values As Collection) As Double
    Dim Entry As DataEntry
    Dim Result As Double
    Dim IsFirst As Boolean

    For Each Entry In Values
        If Not IsFirst Then
            Result = CDbl(Entry.Value)
            IsFirst = True
        ElseIf CDbl(Entry.Value) > Result Then
            Result = CDbl(Entry.Value)
        End If
    Next Entry

    QueryMax = Result
End Function

Public Function QueryProduct(ByVal Values As Collection) As Double
    Dim Entry As DataEntry
    Dim Result As Double

    Result = 1

    For Each Entry In Values
        Result = Result * CDbl(Entry.Value)
    Next Entry

    QueryProduct = Result
End Function