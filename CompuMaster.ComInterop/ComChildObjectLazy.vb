Public Class ComChildObjectLazy(Of P As ComObjectBase)
    Inherits ComRootObjectLazy

    Public Sub New(parentItem As P, obj As Object)
        Me.New(parentItem, parentItem, obj)
    End Sub

    Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, parentItem As P, obj As Object)
        MyBase.New(parentItemResponsibleForDisposal, obj)
        Me.Parent = parentItem
    End Sub

    Public Sub New(parentItemResponsibleForDisposal As ComObjectBase, obj As Object,
                   onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(parentItemResponsibleForDisposal, obj, onDisposeChildrenAction, onClosingAction, onClosedAction)
    End Sub

    Public ReadOnly Property Parent As P

End Class
