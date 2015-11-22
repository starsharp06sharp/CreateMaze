Module AstarModule
    Dim height, width As Integer
    Dim nowQueue As SortedSet(Of Node)
    Dim nowSet As HashSet(Of Integer)
    Dim outSet As HashSet(Of Integer)

    Private Function addNode(cell As Integer, ByRef prev As Node) As Boolean
        If nowSet.Contains(cell) Then
            Dim tmpNode = New Node(cell, prev.cost + 1, prev)
            Dim nodeInQueue = getNode(cell)
            If nodeInQueue.expectCost > tmpNode.expectCost Then
                nowQueue.Remove(nodeInQueue)
                nowQueue.Add(tmpNode)
            End If
        Else
            nowSet.Add(cell)
            nowQueue.Add(New Node(cell, prev.cost + 1, prev))
        End If
        If cell = UnionFind.numOfCell Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function getNode(cell As Integer) As Node
        For Each node In nowQueue
            If node.NO = cell Then
                Return node
            End If
        Next
        Return Nothing
    End Function

    Private Function expand(ByVal node As Node) As Boolean
        Dim cell = node.NO
        Dim adjacentCell As Integer
        adjacentCell = cell - width '上
        If adjacentCell >= 1 AndAlso Not outSet.Contains(adjacentCell) AndAlso Not Form1.hasBorder(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        adjacentCell = cell - 1 '左
        If adjacentCell Mod width <> 0 AndAlso Not outSet.Contains(adjacentCell) AndAlso Not Form1.hasBorder(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        adjacentCell = cell + width '下
        If adjacentCell <= numOfCell AndAlso Not outSet.Contains(adjacentCell) AndAlso Not Form1.hasBorder(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        adjacentCell = cell + 1 '右
        If cell Mod width <> 0 AndAlso Not outSet.Contains(adjacentCell) AndAlso Not Form1.hasBorder(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        Return False
    End Function

    Public Sub findRoute(height As Integer, width As Integer)
        AstarModule.height = height
        AstarModule.width = width
        nowQueue = New SortedSet(Of Node)()
        nowSet = New HashSet(Of Integer)
        nowQueue.Add(New Node(1, 0, Nothing))
        nowSet.Add(1)
        outSet = New HashSet(Of Integer)()

        Dim loopCounter = 0 '临时
        Do
            Dim tmpNode = nowQueue.First()
            outSet.Add(tmpNode.NO)
            If expand(tmpNode) Then
                paintCells(tmpNode)
                Exit Do
            End If
            '标记颜色
            'With Form1.getCell(tmpNode.NO).Interior
            '    .Pattern = Excel.XlPattern.xlPatternSolid
            '    .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            '    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            '    .TintAndShade = 0
            '    .PatternTintAndShade = 0
            'End With

            nowSet.Remove(tmpNode.NO)
            nowQueue.Remove(tmpNode)
            loopCounter += 1
        Loop
    End Sub

    Private Sub paintCells(ByVal node As Node)
        If node.NO = 1 Then
            paintCell(1)
            Exit Sub
        End If
        paintCells(node.prev)
        paintCell(node.NO)
    End Sub

    Private Sub paintCell(cell As Integer)
        With Form1.getCell(cell).Interior
            .Pattern = Excel.XlPattern.xlPatternSolid
            .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
            .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent5
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End Sub
End Module
