Imports System.ComponentModel

Public Class Form1
    Public width, height As Integer
    Private objExcel As New Excel.Application

    Private Function checkInput(ByRef text As String) As Boolean
        If Not IsNumeric(text) Then
            MessageBox.Show("输入不是数字！", "别闹", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If
        If text.IndexOf(".") <> -1 Then
            MessageBox.Show("请勿输入小数！", "别闹", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If
        If Convert.ToInt64(text) > Int32.MaxValue Then
            MessageBox.Show("输入数字过大！", "别闹", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If
        If Convert.ToInt64(text) < Int32.MinValue Then
            MessageBox.Show("输入数字过小！", "别闹", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return False
        End If

        Return True
    End Function

    Public Function hasBorder(cell1 As Integer, cell2 As Integer) As Boolean
        If cell2 = cell1 - width Then
            '2在1上方
            Return IIf(Excel.XlLineStyle.xlLineStyleNone = getCell(cell1).Borders(3).LineStyle, False, True)
        ElseIf cell2 = cell1 - 1 Then
            '2在1左方
            Return IIf(Excel.XlLineStyle.xlLineStyleNone = getCell(cell1).Borders(1).LineStyle, False, True)
        ElseIf cell2 = cell1 + width Then
            '2在1下方
            Return IIf(Excel.XlLineStyle.xlLineStyleNone = getCell(cell1).Borders(4).LineStyle, False, True)
        ElseIf cell2 = cell1 + 1 Then
            '2在1右方
            Return IIf(Excel.XlLineStyle.xlLineStyleNone = getCell(cell1).Borders(2).LineStyle, False, True)
        End If
    End Function

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        objExcel.Quit()
        objExcel = Nothing
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        objExcel.Visible = True
        objExcel.Workbooks.Add()
        objExcel.Caption = "生成迷宫"
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        If Not checkInput(txtWidth.Text) Or Not checkInput(txtHeight.Text) Then
            Exit Sub
        End If
        width = CInt(txtWidth.Text)
        height = CInt(txtHeight.Text)
        initMap()
        Dim rand As New Random()
        Dim cell1, cell2 As Integer
        Do Until UnionFind.isSame(1, numOfCell)
            Do
                cell1 = rand.Next(1, numOfCell)
                cell2 = checkAround(cell1)
            Loop Until cell2 > 0
            UnionFind.union(cell1, cell2)
            deleteBorder(cell1, cell2)
        Loop
        MessageBox.Show("完成！")
    End Sub

    Public Function getCell(ByVal cellNO As Integer) As Excel.Range
        Dim x = cellNO Mod width, y = cellNO \ width + 1
        If x = 0 Then
            x = width
            y -= 1
        End If
        Return objExcel.Cells(y + 1, x + 1)
    End Function

    Private Sub deleteBorder(ByVal cell1 As Integer, ByVal cell2 As Integer)
        If cell2 = cell1 - width Then
            '2在1上方
            getCell(cell1).Borders(3).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            getCell(cell2).Borders(4).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        ElseIf cell2 = cell1 - 1 Then
            '2在1左方
            getCell(cell1).Borders(1).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            getCell(cell2).Borders(2).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        ElseIf cell2 = cell1 + width Then
            '2在1下方
            getCell(cell1).Borders(4).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            getCell(cell2).Borders(3).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        ElseIf cell2 = cell1 + 1 Then
            '2在1右方
            getCell(cell1).Borders(2).LineStyle = Excel.XlLineStyle.xlLineStyleNone
            getCell(cell2).Borders(1).LineStyle = Excel.XlLineStyle.xlLineStyleNone
        End If
    End Sub

    Private Function checkAround(ByVal cell As Integer) As Integer
        Dim adjacentCell As Integer
        adjacentCell = cell - width '上
        If adjacentCell >= 1 AndAlso Not UnionFind.isSame(cell, adjacentCell) Then
            Return adjacentCell
        End If
        adjacentCell = cell - 1 '左
        If adjacentCell Mod width <> 0 AndAlso Not UnionFind.isSame(cell, adjacentCell) Then
            Return adjacentCell
        End If
        adjacentCell = cell + width '下
        If adjacentCell <= numOfCell AndAlso Not UnionFind.isSame(cell, adjacentCell) Then
            Return adjacentCell
        End If
        adjacentCell = cell + 1 '右
        If cell Mod width <> 0 AndAlso Not UnionFind.isSame(cell, adjacentCell) Then
            Return adjacentCell
        End If
        Return 0
    End Function

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        AstarModule.findRoute(height, width)
        MessageBox.Show("完成！")
    End Sub

    Private Sub initMap()
        '设定列宽
        objExcel.Range(objExcel.Columns(2), objExcel.Columns(width + 1)).Select()
        objExcel.Selection.columnwidth = 2
        '设定边框
        objExcel.Range(objExcel.Cells(2, 2), objExcel.Cells(height + 1, width + 1)).Select()
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = 1
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = 1
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = 1
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = 1
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = 1
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = 1
            .Weight = 3
        End With
        UnionFind.init(height * width)
    End Sub
End Class
