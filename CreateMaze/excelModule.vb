Module excelModule
    Private height, width As Integer
    Private objExcel As Excel.Application

    Public Sub init()
        objExcel = New Excel.Application
        objExcel.Visible = True
        objExcel.Workbooks.Add()
        objExcel.Caption = "生成迷宫"
    End Sub

    Public Sub dispose()
        objExcel.Quit()
        objExcel = Nothing
    End Sub

    Public Sub initMap(height As Integer, width As Integer)
        excelModule.height = height
        excelModule.width = width
        '设定列宽
        objExcel.Range(objExcel.Columns(2), objExcel.Columns(width + 1)).Select()
        objExcel.Selection.columnwidth = 2
        '设定边框
        objExcel.Range(objExcel.Cells(2, 2), objExcel.Cells(height + 1, width + 1)).Select()
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = 3
        End With
        With objExcel.Selection.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = 3
        End With
        UnionFind.init(height * width)
    End Sub

    Public Function hasBorderBetween(cell1 As Integer, cell2 As Integer) As Boolean
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

    Public Sub deleteBorder(ByVal cell1 As Integer, ByVal cell2 As Integer)
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


    Public Function getCell(ByVal cellNO As Integer) As Excel.Range
        Dim x = cellNO Mod width, y = cellNO \ width + 1
        If x = 0 Then
            x = width
            y -= 1
        End If
        Return objExcel.Cells(y + 1, x + 1)
    End Function
End Module
