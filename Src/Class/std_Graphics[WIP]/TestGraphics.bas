Attribute VB_Name = "TestGraphics"

Option Explicit

Public Function FonctionOpenGL()

    Dim Positions() As Double

    ReDim Positions(39)

    Positions(00) = -1.0: Positions(01) =  -1.0: Positions(02) =  -1.0: Positions(03) =  0.0: Positions(04) =  0.0
    Positions(05) = +1.0: Positions(06) =  -1.0: Positions(07) =  -1.0: Positions(08) =  1.0: Positions(09) =  0.0
    Positions(10) = +1.0: Positions(11) =  +1.0: Positions(12) =  -1.0: Positions(13) =  1.0: Positions(14) =  1.0
    Positions(15) = -1.0: Positions(16) =  +1.0: Positions(17) =  -1.0: Positions(18) =  0.0: Positions(19) =  1.0
    Positions(20) = -1.0: Positions(21) =  -1.0: Positions(22) =  +1.0: Positions(23) =  0.0: Positions(24) =  0.0
    Positions(25) = +1.0: Positions(26) =  -1.0: Positions(27) =  +1.0: Positions(28) =  1.0: Positions(29) =  0.0
    Positions(30) = +1.0: Positions(31) =  +1.0: Positions(32) =  +1.0: Positions(33) =  1.0: Positions(34) =  1.0
    Positions(35) = -1.0: Positions(36) =  +1.0: Positions(37) =  +1.0: Positions(38) =  0.0: Positions(39) =  1.0


    Dim Indices() As Long
    ReDim Indices(35)
    'Front
    Indices(0) = 2: Indices(1) = 0: Indices(2) = 1
    Indices(3) = 2: Indices(4) =  3: Indices(5) =  0
    'Back
    Indices(6) = 3: Indices(7) =  4: Indices(8) =  0
    Indices(9) = 3: Indices(10) =  7: Indices(11) =  4
    'Right
    Indices(12) = 6: Indices(13) =  1: Indices(14) =  5
    Indices(15) = 6: Indices(16) =  2: Indices(17) =  1
    'Left
    Indices(18) = 7: Indices(19) =  5: Indices(20) =  4
    Indices(21) = 7: Indices(22) =  6: Indices(23) =  5
    'Up
    Indices(24) = 6: Indices(25) =  3: Indices(26) =  2
    Indices(27) = 6: Indices(28) =  7: Indices(29) =  3
    'Down
    Indices(30) = 1: Indices(31) =  4: Indices(32) =  5
    Indices(33) = 1: Indices(34) =  0: Indices(35) =  4

    



    If LoadLibrary("C:\Users\deallulic\Documents\Freeglut\freeglut64.dll") = 0 Then
        MsgBox "Cant load Library"
        Exit Function
    End If
    glutInit 0&, ""
    glutInitDisplayMode GLUT_RGBA Or GLUT_DOUBLE Or GLUT_DEPTH

    glutCreateWindow "Test Cube"
    glutSetOption GLUT_ACTION_ON_WINDOW_CLOSE, GLUT_ACTION_GLUTMAINLOOP_RETURNS

    Dim VB As std_VertexBuffer
    Set VB = std_VertexBuffer.Create(Positions)

    Dim Shader As std_Shader
    Set Shader = std_Shader.CreateFromFile("C:\Users\deallulic\Documents\GitHub\AL_StdLib\Src\Class\std_Graphics[WIP]\Vertex.Shader", "C:\Users\deallulic\Documents\GitHub\AL_StdLib\Src\Class\std_Graphics[WIP]\Fragment.Shader")

    
    glutDisplayFunc AddressOf CallBackDraw
    glutMainLoop
End Function

' Fonction d'affichage
Public Sub CallBackDraw()
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glutSwapBuffers
End Sub




