Imports System.ComponentModel

Public Class Form1
    Private width, height As Integer

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        excelModule.dispose()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        excelModule.init()
    End Sub

    Private Sub btnCreate_Click(sender As Object, e As EventArgs) Handles btnCreate.Click
        If Not checkInput(txtWidth.Text) Or Not checkInput(txtHeight.Text) Then
            Exit Sub
        End If
        width = CInt(txtWidth.Text)
        height = CInt(txtHeight.Text)
        initMap(height, width)
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

    Private Sub btnFind_Click(sender As Object, e As EventArgs) Handles btnFind.Click
        AstarModule.findRoute(height, width)
        MessageBox.Show("完成！")
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
End Class
