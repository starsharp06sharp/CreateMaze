Module AstarModule
    Private height, width As Integer
    Private nowQueue As SortedSet(Of Node)
    Private nowSet As HashSet(Of Integer)
    Private outSet As HashSet(Of Integer)

    Public Sub findRoute(height As Integer, width As Integer)
        AstarModule.height = height
        AstarModule.width = width
        nowQueue = New SortedSet(Of Node)()
        nowSet = New HashSet(Of Integer)
        nowQueue.Add(New Node(1, 0, Nothing))
        nowSet.Add(1)
        outSet = New HashSet(Of Integer)()

        Do
            Dim tmpNode = nowQueue.First()
            outSet.Add(tmpNode.NO)
            If expand(tmpNode) Then
                paintCells(tmpNode)
                paintCell(height * width, "dark")
                Exit Do
            End If
            paintCell(tmpNode.NO, "light")

            nowSet.Remove(tmpNode.NO)
            nowQueue.Remove(tmpNode)
        Loop
    End Sub

    Private Function expand(ByVal node As Node) As Boolean
        Dim cell = node.NO
        Dim adjacentCell As Integer
        adjacentCell = cell - width '上
        If adjacentCell >= 1 AndAlso Not outSet.Contains(adjacentCell) AndAlso Not hasBorderBetween(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        adjacentCell = cell - 1 '左
        If adjacentCell Mod width <> 0 AndAlso Not outSet.Contains(adjacentCell) AndAlso Not hasBorderBetween(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        adjacentCell = cell + width '下
        If adjacentCell <= numOfCell AndAlso Not outSet.Contains(adjacentCell) AndAlso Not hasBorderBetween(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        adjacentCell = cell + 1 '右
        If cell Mod width <> 0 AndAlso Not outSet.Contains(adjacentCell) AndAlso Not hasBorderBetween(cell, adjacentCell) Then
            If addNode(adjacentCell, node) Then Return True
        End If
        Return False
    End Function

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

    Private Sub paintCells(ByVal node As Node)
        If node.NO = 1 Then
            paintCell(1, "dark")
            Exit Sub
        End If
        paintCells(node.prev)
        paintCell(node.NO, "dark")
    End Sub

    Private Sub paintCell(cell As Integer, shade As String)
        If shade = "dark" Then
            excelModule.getCell(cell).FormulaR1C1 = "*"
        ElseIf shade = "light"
            If Form1.chkShowProcess.Checked Then
                With excelModule.getCell(cell).Interior
                    .Pattern = Excel.XlPattern.xlPatternSolid
                    .PatternColorIndex = Excel.XlPattern.xlPatternAutomatic
                    .ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Else
            Throw New ArgumentException("Unknown shade argument")
        End If
    End Sub
End Module
