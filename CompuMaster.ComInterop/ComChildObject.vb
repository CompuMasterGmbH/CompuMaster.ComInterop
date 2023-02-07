Namespace CompuMaster.ComInterop

    ''' <summary>
    ''' Wrapper for COM child node, typically a document, sheet, module, etc.
    ''' </summary>
    ''' <typeparam name="TParentWrapper">A type of the wrapper for the parent node</typeparam>
    ''' <typeparam name="TComObject">A type of the child node (usually used with COM interop assemblies) or just a System.Object</typeparam>
    Public Class ComChildObject(Of TParentWrapper As ComObjectBase, TComObject As Class)
        Inherits ComRootObject(Of TComObject)

        ''' <summary>
        ''' Create a new wrapper for a COM child object
        ''' </summary>
        ''' <param name="parentItem">A parent item which is referenced in property <see cref="Parent"/> and which triggers disposal of this item</param>
        ''' <param name="obj"></param>
        Public Sub New(parentItem As TParentWrapper, obj As TComObject)
            Me.New(parentItem, parentItem, obj)
        End Sub

        ''' <summary>
        ''' Create a new wrapper for a COM child object
        ''' </summary>
        ''' <param name="parentItemResponsibleForDisposal">A parent item which triggers disposal of this item</param>
        ''' <param name="parentItem">A parent item which is referenced in property <see cref="Parent"/></param>
        ''' <param name="obj"></param>
        Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, parentItem As TParentWrapper, obj As TComObject)
            MyBase.New(parentItemResponsibleForDisposal, obj, Nothing, Nothing, Nothing)
            Me.Parent = parentItem
        End Sub

        ''' <summary>
        ''' Create a new wrapper for a COM child object
        ''' </summary>
        ''' <param name="parentItem">A parent item which is referenced in property <see cref="Parent"/> and which triggers disposal of this item</param>
        ''' <param name="obj"></param>
        ''' <param name="onDisposeChildrenAction"></param>
        ''' <param name="onClosingAction"></param>
        ''' <param name="onClosedAction"></param>
        Public Sub New(parentItem As ComObjectBase, obj As TComObject,
                   onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
            MyBase.New(parentItem, obj, onDisposeChildrenAction, onClosingAction, onClosedAction)
        End Sub

        ''' <summary>
        ''' The parent wrapper object
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Parent As TParentWrapper

    End Class

End Namespace