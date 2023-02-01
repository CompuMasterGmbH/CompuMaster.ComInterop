Public MustInherit Class ObjectReadOnlyCollectionBase(Of TCollectionComObject As Class,
                                                          TCollectionWrapper As ComChildObject(Of TCollectionComObject, TParentWrapper),
                                                          TParentWrapper As ComObjectBase,
                                                          TChildComObject As Class,
                                                          TChildWrapper As ComChildObject(Of TChildComObject, TCollectionWrapper)
                                                          )
    Inherits ComChildObject(Of TCollectionComObject, TParentWrapper)

    Public Sub New(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TCollectionComObject)
        MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
    End Sub

    Public MustOverride ReadOnly Property Count As Integer

    Public MustOverride ReadOnly Property Item(index As Integer) As TChildWrapper

    'Protected Shared Function CreateChildInstance(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TCollectionComObject) As TChildWrapper

End Class
