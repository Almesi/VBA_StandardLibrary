Attribute VB_Name = "VBFPExamples"

Option Explicit

Public Sub ExampleArithmetic()
    Debug.Print "=== Arithmetic ==="

    Call VBFPPrintExtra("10 + 5 = ", VBFPBoundF(VBFPNamedF("Add")     , VBFPValue(10)).Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 - 5 = ", VBFPBoundF(VBFPNamedF("Subtract"), VBFPValue(10)).Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 * 5 = ", VBFPBoundF(VBFPNamedF("Multiply"), VBFPValue(10)).Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 / 5 = ", VBFPBoundF(VBFPNamedF("Divide")  , VBFPValue(10)).Run(VBFPValue(5)))
End Sub

Public Sub ExampleBinding()
    Debug.Print "=== Binding ==="

    Dim Add10 As IFunction
    Set Add10 = VBFPBoundF(VBFPNamedF("Add"), VBFPValue(10))

    Call VBFPPrintExtra("10 + 1 = " , Add10.Run(VBFPValue(1)))
    Call VBFPPrintExtra("10 + 5 = " , Add10.Run(VBFPValue(5)))
    Call VBFPPrintExtra("10 + 20 = ", Add10.Run(VBFPValue(20)))
End Sub

Public Sub ExampleMap()
    Debug.Print "=== Map ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction : Set Elements = VBFPValue(Data)
    Dim Add10    As IFunction : Set Add10    = VBFPBoundF(VBFPNamedF("Add"), VBFPValue(10))
    Dim Result   As IFunction : Set Result   = VBFPBoundF(VBFPNamedF("Map"), Elements).Run(Add10)

    Call VBFPPrint(Result)

End Sub

Public Sub ExampleMapArithmetic()
    Debug.Print "=== Map Arithmetic ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Dim Func As IFunction
    Dim Result As IFunction

    Set Func   = VBFPBoundF(VBFPNamedF("Add"), VBFPValue(2))
    Set Result = VBFPBoundF(VBFPNamedF("Map"), Elements).Run(Func)
    Debug.Print "Add 2:"
    Call VBFPPrint(Result)

    Set Func   = VBFPBoundF(VBFPNamedF("Multiply"), VBFPValue(10))
    Set Result = VBFPBoundF(VBFPNamedF("Map"), Elements).Run(Func)
    Debug.Print "Multiply 10:"
    Call VBFPPrint(Result)

    Set Func   = VBFPBoundF(VBFPNamedF("Subtract"), VBFPValue(1))
    Set Result = VBFPBoundF(VBFPNamedF("Map"), Elements).Run(Func)
    Debug.Print "Subtract 1:"
    Call VBFPPrint(Result)
End Sub

Public Sub ExamplePredicates()
    Debug.Print "=== Predicates ==="

    Call VBFPPrintExtra("5 > 2 = ", VBFPBoundF(VBFPNamedF("GreaterThan"), VBFPValue(5)).Run(VBFPValue(2)))
    Call VBFPPrintExtra("5 < 2 = ", VBFPBoundF(VBFPNamedF("LessThan"), VBFPValue(5)).Run(VBFPValue(2)))
    Call VBFPPrintExtra("5 = 5 = ", VBFPBoundF(VBFPNamedF("Equal"), VBFPValue(5)).Run(VBFPValue(5)))
End Sub

Public Sub ExampleStrings()
    Debug.Print "=== Strings ==="

    Dim Hello As IFunction
    Dim World As IFunction

    Set Hello = VBFPValue("Hello, ")
    Set World = VBFPValue("World!")

    Call VBFPPrintExtra("Concatenate", VBFPBoundF(VBFPNamedF("Concatenate"), Hello).Run(World))
    Call VBFPPrintExtra("UpperCase"  , VBFPNamedF("UpperCase").Run(VBFPValue("hello world")))
End Sub

Public Sub ExampleArrayFunctions()
    Debug.Print "=== Array Functions ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(10)
    Set Data(1) = VBFPValue(20)
    Set Data(2) = VBFPValue(30)
    Set Data(3) = VBFPValue(40)
    Set Data(4) = VBFPValue(50)

    Dim Elements As IFunction
    Set Elements = VBFPValue(Data)

    Call VBFPPrintExtra("First  = ", VBFPNamedF("First").Run(Elements))
    Call VBFPPrintExtra("Last   = ", VBFPNamedF("Last").Run(Elements))
    Call VBFPPrintExtra("Length = ", VBFPNamedF("ArrayLength").Run(Elements))
End Sub

Public Sub ExampleFold()
    Debug.Print "=== Fold ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction: Set Elements = VBFPValue(Data)
    Dim Add      As IFunction: Set Add      = VBFPNamedF("Add")
    Dim Initial  As IFunction: Set Initial  = VBFPValue(0)
    Dim Result   As IFunction: Set Result   = VBFPBoundF(VBFPNamedF("Fold"), Elements, Add).Run(Initial)
    Call VBFPPrintExtra("1 + 2 + 3 + 4 + 5 = ", Result)
End Sub

Public Sub ExampleFactorial()
    Debug.Print "=== Fold / Factorial ==="

    Dim Data(4) As IFunction

    Set Data(0) = VBFPValue(1)
    Set Data(1) = VBFPValue(2)
    Set Data(2) = VBFPValue(3)
    Set Data(3) = VBFPValue(4)
    Set Data(4) = VBFPValue(5)

    Dim Elements As IFunction : Set Elements = VBFPValue(Data)
    Dim Multiply As IFunction : Set Multiply = VBFPNamedF("Multiply")
    Dim Initial  As IFunction : Set Initial  = VBFPValue(1)
    Dim Result   As IFunction : Set Result   = VBFPBoundF(VBFPNamedF("Fold"), Elements, Multiply).Run(Initial)
    Call VBFPPrintExtra("1 * 2 * 3 * 4 * 5 = ", Result)
End Sub

Public Sub ExampleIfThenElse()
    Debug.Print "=== IfThenElse ==="

    Dim Condition As IFunction : Set Condition = VBFPValue(True)
    Dim WhenTrue  As IFunction : Set WhenTrue  = VBFPValue("Condition was true")
    Dim WhenFalse As IFunction : Set WhenFalse = VBFPValue("Condition was false")
    Dim Result    As IFunction : Set Result    = VBFPBoundF(VBFPNamedF("IfThenElse"), Condition, WhenTrue).Run(WhenFalse)
    Call VBFPPrint(Result)
End Sub

Public Sub ExampleCompose()
    Debug.Print "=== Compose ==="

    Dim Value     As IFunction : Set Value     = VBFPValue(5)
    Dim Add10     As IFunction : Set Add10     = VBFPBoundF(VBFPNamedF("Add"), VBFPValue(10))
    Dim Multiply2 As IFunction : Set Multiply2 = VBFPBoundF(VBFPNamedF("Multiply"), VBFPValue(2))
    Dim Composed  As IFunction : Set Composed  = VBFPComposedF(Add10, Multiply2)
    Dim Result    As IFunction : Set Result    = Composed.Run(Value)
    Call VBFPPrintExtra("(5 * 2) + 10 = ", Result)
End Sub

Public Sub ExamplePipeline()
    Debug.Print "=== Pipeline ==="

    Dim Value     As IFunction : Set Value     = VBFPValue(5)
    Dim Add10     As IFunction : Set Add10     = VBFPBoundF(VBFPNamedF("Add"), VBFPValue(10))
    Dim Multiply2 As IFunction : Set Multiply2 = VBFPBoundF(VBFPNamedF("Multiply"), VBFPValue(2))
    Dim Result    As IFunction : Set Result    = Pipe(Value, Add10, Multiply2)
    Call VBFPPrintExtra("(5 + 10) * 2 = ", Result)
End Sub

Public Sub ExampleCurry()
    Debug.Print "=== Curry, yummy ==="

    Dim Value     As IFunction: Set Value = VBFPValue(5)
    Dim Curried   As IFunction: Set Curried = VBFPCurriedF(VBFPNamedF("Add"), VBFPValue(3))
    Dim Result    As IFunction: Set Result = Curried.Run(VBFPValue(10))
    Call VBFPPrintExtra("3 + 10 = ", Result)
End Sub

Public Sub ExampleMemoized()
    Debug.Print "=== Memoized ==="

    Dim Func      As IFunction: Set Func = VBFPMemoizedF(VBFPNamed("Factorial"))
    Call VBFPPrintExtra("1! = ", Func.Run(VBFPValue(1)))
    Call VBFPPrintExtra("2! = ", Func.Run(VBFPValue(2)))
    Call VBFPPrintExtra("3! = ", Func.Run(VBFPValue(3)))
    Call VBFPPrintExtra("4! = ", Func.Run(VBFPValue(4)))
    Call VBFPPrintExtra("5! = ", Func.Run(VBFPValue(5)))
    Call VBFPPrintExtra("6! = ", Func.Run(VBFPValue(6)))
    Call VBFPPrintExtra("5! = ", Func.Run(VBFPValue(5)))

    Set Func = VBFPMemoizedF(VBFPNamed("Multiply"))
    Call VBFPPrintExtra("1*1 = ", Func.Run(VBFPValue(1), VBFPValue(1)))
    Call VBFPPrintExtra("2*1 = ", Func.Run(VBFPValue(2), VBFPValue(1)))
    Call VBFPPrintExtra("3*1 = ", Func.Run(VBFPValue(3), VBFPValue(1)))
    Call VBFPPrintExtra("1*2 = ", Func.Run(VBFPValue(1), VBFPValue(2)))
    Call VBFPPrintExtra("2*2 = ", Func.Run(VBFPValue(2), VBFPValue(2)))
    Call VBFPPrintExtra("3*2 = ", Func.Run(VBFPValue(3), VBFPValue(2)))
    Call VBFPPrintExtra("2*2 = ", Func.Run(VBFPValue(2), VBFPValue(2)))
End Sub

Public Sub ExampleNormalFunctions()
    Debug.Print "=== Normal Function ==="

    Dim Func      As IFunction: Set Func = NamedFunction.Create("NormalAdd", False, False)
    Call VBFPPrintExtra("3 = ", Func.Run(1, 2))

End Sub


Public Sub RunExamples()
    ExampleArithmetic
    ExampleBinding
    ExampleMap
    ExampleMapArithmetic
    ExamplePredicates
    ExampleStrings
    ExampleArrayFunctions
    ExampleFold
    ExampleFactorial
    ExampleIfThenElse
    ExampleCompose
    ExamplePipeline
    ExampleMemoized
    ExampleCurry
    ExampleNormalFunctions
End Sub