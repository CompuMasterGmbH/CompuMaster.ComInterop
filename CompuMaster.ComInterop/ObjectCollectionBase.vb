Public MustInherit Class ObjectCollectionBase(Of TCollectionComObject As Class, TCollectionWrapper As ComChildObject(Of TCollectionComObject, TParentWrapper), TParentWrapper As ComObjectBase, TChildComObject As Class, TChildWrapper As ComInterop.ComChildObject(Of TChildComObject, TCollectionWrapper))
    Inherits ObjectReadOnlyCollectionBase(Of TCollectionComObject, TCollectionWrapper, TParentWrapper, TChildComObject, TChildWrapper)

    Public Sub New(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TCollectionComObject)
        MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
    End Sub

    Public MustOverride Sub Add(item As TChildWrapper)

    Public MustOverride Sub Insert(index As Integer, item As TChildWrapper)

    Public MustOverride Sub Remove(item As TChildWrapper)

    Public MustOverride Sub RemoveAt(index As Integer)

End Class
