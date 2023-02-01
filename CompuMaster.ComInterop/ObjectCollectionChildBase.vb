Public MustInherit Class ObjectCollectionChildBase(Of P As ComObjectBase, T As ComObjectBase)
    Inherits ComObjectBase

    Friend Sub New(c As ObjectCollectionBase(Of P, T), accessModuleComObject As Object)
        MyBase.New(c.Parent, accessModuleComObject)
        Parent = c
    End Sub

    Public ReadOnly Parent As ObjectCollectionBase(Of P, T)

End Class
