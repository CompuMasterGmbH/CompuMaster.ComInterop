''' <summary>
''' Wrapper for COM root node, typically the application
''' </summary>
''' <typeparam name="TComObject">A type of the application (usually used with COM interop assemblies) or System.Object (usually used with COM objects created with CreateObject)</typeparam>
Public Class ComRootObject(Of TComObject)
    Inherits ComObjectBase

    Public Sub New(obj As TComObject)
        MyBase.New(Nothing, obj)
    End Sub

    Public Sub New(obj As TComObject,
                           onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(Nothing, obj)
        Me._OnDisposeChildrenAction = onDisposeChildrenAction
        Me._OnClosingAction = onClosingAction
        Me._OnClosedAction = onClosedAction
    End Sub

    Protected Friend Sub New(parentItemResponsibleForDisposal As Global.CompuMaster.ComInterop.ComObjectBase, obj As TComObject)
        MyBase.New(parentItemResponsibleForDisposal, obj)
    End Sub

    Protected Friend Sub New(parentItemResponsibleForDisposal As Global.CompuMaster.ComInterop.ComObjectBase, obj As TComObject,
                           onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(parentItemResponsibleForDisposal, obj)
        Me._OnDisposeChildrenAction = onDisposeChildrenAction
        Me._OnClosingAction = onClosingAction
        Me._OnClosedAction = onClosedAction
    End Sub

    Private _OnDisposeChildrenAction As OnDisposeChildrenAction
    Public Delegate Sub OnDisposeChildrenAction(instance As ComRootObject(Of TComObject))

    Private _OnClosingAction As OnClosingAction
    Public Delegate Sub OnClosingAction(instance As ComRootObject(Of TComObject))

    Private _OnClosedAction As OnClosedAction
    Public Delegate Sub OnClosedAction(instance As ComRootObject(Of TComObject))

    Protected Overrides Sub OnDisposeChildren()
        If _OnDisposeChildrenAction IsNot Nothing Then _OnDisposeChildrenAction(Me)
    End Sub

    Protected Overrides Sub OnClosing()
        If _OnClosingAction IsNot Nothing Then _OnClosingAction(Me)
    End Sub

    Protected Overrides Sub OnClosed()
        If _OnClosedAction IsNot Nothing Then _OnClosedAction(Me)
    End Sub

    Public ReadOnly Property ComObjectStronglyTyped As TComObject
        Get
            Return CType(MyBase.ComObject, TComObject)
        End Get
    End Property

End Class
