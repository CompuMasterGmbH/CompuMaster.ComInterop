Public MustInherit Class ObjectCollectionBase(Of TParentWrapper As ComObjectBase,
                                                  TCollectionComObject As Class,
                                                  TCollectionWrapper As ComChildObject(Of TParentWrapper, TCollectionComObject),
                                                  TChildComObject As Class,
                                                  TChildWrapper As ComInterop.ComChildObject(Of TCollectionWrapper, TChildComObject)
                                                  )
    Inherits ObjectReadOnlyCollectionBase(Of TParentWrapper, TCollectionComObject, TCollectionWrapper, TChildComObject, TChildWrapper)

    Public Sub New(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TCollectionComObject)
        MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
    End Sub

    Public MustOverride Sub Add(item As TChildWrapper)

    Public MustOverride Sub Insert(index As Integer, item As TChildWrapper)

    Public MustOverride Sub Remove(item As TChildWrapper)

    Public MustOverride Sub RemoveAt(index As Integer)

End Class
