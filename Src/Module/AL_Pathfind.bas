Attribute VB_Name = "AL_Pathfind"

Function AL_Pathfind_AStar(StartNode As AL_Node, GoalNode As AL_Node, ByRef ImpassableNodes() As AL_Node, MaxZ As Integer, MinZ As Integer, MaxY As Integer, MinY As Integer, MaxX As Integer, MinX As Integer) As AL_Node()

    Dim OpenNodes() As New AL_Node
    Dim ClosedNodes() As New AL_Node
    Dim FinalPath() As New AL_Node
    Dim CurrentNode As New AL_Node
    Dim NeighbourNode As New AL_Node
    Dim TempNode As New AL_Node

    Dim TempIndex As Integer
    Dim OpenNodesIndex As Integer
    Dim ClosedNodesIndex As Integer
    Dim ImpassableNodesIndex As Integer
    Dim CurrentNodeIndex As Integer
    
    ' Initialize Indices
    OpenNodesIndex = 0
    ClosedNodesIndex = 0
    ImpassableNodesIndex = 0
    ReDim Preserve OpenNodes(OpenNodesIndex)
    ReDim Preserve ClosedNodes(ClosedNodesIndex)
    ReDim Preserve FinalPath(OpenNodesIndex)

    ' Initialize first Node
    Set OpenNodes(OpenNodesIndex) = StartNode
    OpenNodes(OpenNodesIndex).CalculateDistance OpenNodes(OpenNodesIndex), StartNode, GoalNode
    Set CurrentNode = OpenNodes(OpenNodesIndex)
    CurrentNodeIndex = OpenNodesIndex
    CurrentNode.CalculateDistance CurrentNode, StartNode, GoalNode

    Do Until OpenNodesIndex < 0
        
        ' If on Goal reconstruct Path
        If CurrentNode.IsNode(GoalNode) = True Then
            Set FinalPath(0) = CurrentNode
            Do Until CurrentNode.Parent Is Nothing
                AL_Array_Push_Obj FinalPath, CurrentNode.Parent
                Set CurrentNode = CurrentNode.Parent
                I = I + 1
            Loop
            AL_Pathfind_AStar = FinalPath
            Exit Function
        End If

        ' For each 3-Dimensional Neighbour of CurrentNode (which is in Range)
        For Z = -1 To 1
            If (Z + CurrentNode.Z < MaxZ) And (Z + CurrentNode.Z > MinZ) Then
                For Y = -1 To 1
                    If (Y + CurrentNode.Y < MaxY) And (Y + CurrentNode.Y > MinY) Then
                        For X = -1 To 1
                            If (X + CurrentNode.X < MaxX) And (X + CurrentNode.X > MinX) Then
                                NeighbourNode.LetPoint CurrentNode.X + X, CurrentNode.Y + Y, CurrentNode.Z + Z
                                ' If Neighbour is CurrentNode skip this node
                                If NeighbourNode.IsNode(CurrentNode) = False Then
                                    
                                    ' For each Impassable Node check if its the current one
                                    If Not ImpassableNodes(0) Is Nothing Then
                                        For I = 0 To ImpassableNodesIndex
                                            If NeighbourNode.IsNode(ImpassableNodes(I)) = True Then
                                                GoTo NodeUnavailable
                                            End If
                                        Next
                                    End If

                                    ' For each Closed Node check if its the current one
                                    For I = 0 To ClosedNodesIndex
                                        If NeighbourNode.IsNode(ClosedNodes(I)) = True Then
                                            GoTo NodeUnavailable
                                        End If
                                    Next

                                    ' Give NeighbourNode Values and pushes it into the Array
                                    OpenNodesIndex = OpenNodesIndex + 1
                                    NeighbourNode.CalculateDistance NeighbourNode, StartNode, GoalNode
                                    NeighbourNode.Parent = CurrentNode
                                    AL_Array_Push_Obj OpenNodes, NeighbourNode
                                End If
                                Set NeighbourNode = Nothing
                                Else
NodeUnavailable:
                            End If
                        Next
                    End If
                Next
            End If
        Next

        ' Transfer CurrentNode from OpenNodes to ClosedNodes
        OpenNodesIndex = OpenNodesIndex - 1
        If OpenNodesIndex < 0 Then
            Exit Function
        End If
        AL_Array_Delete_Obj OpenNodes, CurrentNodeIndex
        AL_Array_Push_Obj ClosedNodes, CurrentNode
        ClosedNodesIndex = ClosedNodesIndex + 1
        
        ' Find cheapest Node
        Set TempNode = CurrentNode
        For I = 0 To OpenNodesIndex
            If OpenNodes(I).G_Cost < TempNode.G_Cost Then
                Set TempNode = OpenNodes(I)
                TempIndex = I
            End If
        Next
        Set CurrentNode = TempNode
        CurrentNodeIndex = TempIndex
    Loop

End Function