Namespace CompuMaster.ComInterop

    ''' <summary>
    ''' Safe design pattern and implementation for read-only collections of COM objects
    ''' </summary>
    ''' <typeparam name="TParentWrapper"></typeparam>
    ''' <typeparam name="TCollectionComObject"></typeparam>
    ''' <typeparam name="TCollectionWrapper"></typeparam>
    ''' <typeparam name="TChildComObject"></typeparam>
    ''' <typeparam name="TChildWrapper"></typeparam>
    Public MustInherit Class ObjectReadOnlyCollectionBase(Of TParentWrapper As ComObjectBase,
                                                              TCollectionComObject As Class,
                                                              TCollectionWrapper As ComChildObject(Of TParentWrapper, TCollectionComObject),
                                                              TChildComObject As Class,
                                                              TChildWrapper As ComChildObject(Of TCollectionWrapper, TChildComObject)
                                                              )
        Inherits ComChildObject(Of TParentWrapper, TCollectionComObject)

        Public Sub New(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TCollectionComObject)
            MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
        End Sub

        Public MustOverride ReadOnly Property Count As Integer

        Public MustOverride ReadOnly Property Item(index As Integer) As TChildWrapper

        'Protected Shared Function CreateChildInstance(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TCollectionComObject) As TChildWrapper

    End Class

End Namespace