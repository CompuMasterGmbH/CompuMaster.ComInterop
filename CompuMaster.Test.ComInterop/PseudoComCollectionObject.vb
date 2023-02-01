Public Class PseudoComCollectionObject
    Public Sub New(length As Integer)
        Count = length
    End Sub

    Public ReadOnly Property Item(index As Integer) As PseudoComItemObject
        Get
            Return New PseudoComItemObject("List item " & index + 1)
        End Get
    End Property

    Public Property Count As Integer

End Class