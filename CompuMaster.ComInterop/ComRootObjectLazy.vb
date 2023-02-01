Public Class ComRootObjectLazy
    Inherits ComObjectBase

    Public Sub New(obj As Object)
        MyBase.New(Nothing, obj)
    End Sub

    Public Sub New(obj As Object,
                   onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(Nothing, obj)
        Me._OnDisposeChildrenAction = onDisposeChildrenAction
        Me._OnClosingAction = onClosingAction
        Me._OnClosedAction = onClosedAction
    End Sub

    Protected Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, obj As Object)
        MyBase.New(parentItemResponsibleForDisposal, obj)
    End Sub

    Protected Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, obj As Object,
                   onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(parentItemResponsibleForDisposal, obj)
        Me._OnDisposeChildrenAction = onDisposeChildrenAction
        Me._OnClosingAction = onClosingAction
        Me._OnClosedAction = onClosedAction
    End Sub

    Private _OnDisposeChildrenAction As OnDisposeChildrenAction
    Public Delegate Sub OnDisposeChildrenAction(instance As ComRootObjectLazy)

    Private _OnClosingAction As OnClosingAction
    Public Delegate Sub OnClosingAction(instance As ComRootObjectLazy)

    Private _OnClosedAction As OnClosedAction
    Public Delegate Sub OnClosedAction(instance As ComRootObjectLazy)

    Protected Overrides Sub OnDisposeChildren()
        If _OnDisposeChildrenAction IsNot Nothing Then _OnDisposeChildrenAction(Me)
    End Sub

    Protected Overrides Sub OnClosing()
        If _OnClosingAction IsNot Nothing Then _OnClosingAction(Me)
    End Sub

    Protected Overrides Sub OnClosed()
        If _OnClosedAction IsNot Nothing Then _OnClosedAction(Me)
    End Sub

End Class
