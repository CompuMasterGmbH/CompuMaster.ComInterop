Namespace CompuMaster.ComInterop

    ''' <summary>
    ''' Wrapper for COM child node, typically a document, sheet, module, etc.
    ''' </summary>
    ''' <typeparam name="TParentWrapper">A type of the wrapper for the parent node</typeparam>
    ''' <typeparam name="TComObject">A type of the child node (usually used with COM interop assemblies) or just a System.Object</typeparam>
    Public Class ComChildObject(Of TParentWrapper As ComObjectBase, TComObject As Class)
        Inherits ComRootObject(Of TComObject)

        Public Sub New(parentItem As TParentWrapper, obj As TComObject)
            Me.New(parentItem, parentItem, obj)
        End Sub

        Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, parentItem As TParentWrapper, obj As TComObject)
            MyBase.New(parentItemResponsibleForDisposal, obj, Nothing, Nothing, Nothing)
            Me.Parent = parentItem
        End Sub

        Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, obj As TComObject,
                   onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
            MyBase.New(parentItemResponsibleForDisposal, obj, onDisposeChildrenAction, onClosingAction, onClosedAction)
        End Sub

        Public ReadOnly Property Parent As TParentWrapper

    End Class

End Namespace