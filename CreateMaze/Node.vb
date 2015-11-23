Public Class Node
    Implements IComparable(Of Node)
    Public x, y As Integer
    Public NO As Integer
    Public prev As Node
    Public cost As Integer
    Private manhattanExpect As Integer

    Public ReadOnly Property expectCost As Integer
        Get
            Return cost + manhattanExpect
        End Get
    End Property

    Public Sub New(ByVal NO As Integer, ByVal cost As Integer, ByRef prev As Node)
        Me.NO = NO
        Dim x = NO Mod Form1.width, y = NO \ Form1.width + 1
        If x = 0 Then
            x = Form1.width
            y -= 1
        End If
        Me.x = x
        Me.y = y
        Me.cost = cost
        Me.prev = prev
        Me.manhattanExpect = Form1.width - x + Form1.height - y
        Me.manhattanExpect *= 1.5    '估测常数，可调整
    End Sub

    Public Function CompareTo(otherNode As Node) As Integer Implements IComparable(Of Node).CompareTo
        If Me.expectCost = otherNode.expectCost Then
            Return Me.NO.CompareTo(otherNode.NO)
        End If
        Return Me.expectCost.CompareTo(otherNode.expectCost)
    End Function
End Class
