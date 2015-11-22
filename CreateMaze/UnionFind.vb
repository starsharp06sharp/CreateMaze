Module UnionFind
    Public numOfCell As Integer
    Private father() As Integer

    Public Sub init(ByVal n As Integer)
        numOfCell = n
        father = New Integer(n) {}
        For i = 1 To n
            father(i) = i
        Next
    End Sub

    Public Function getFather(ByVal n As Integer) As Integer
        If father(n) = n Then
            Return n
        Else
            father(n) = getFather(father(n))
            Return father(n)
        End If
    End Function

    Public Function isSame(ByVal a As Integer, ByVal b As Integer) As Boolean
        Dim fa = getFather(a), fb = getFather(b)
        Return IIf(fa = fb, True, False)
    End Function

    Public Sub union(ByVal a As Integer, ByVal b As Integer)
        Dim fa = getFather(a), fb = getFather(b)
        If fa <> fb Then father(fa) = fb
    End Sub
End Module
