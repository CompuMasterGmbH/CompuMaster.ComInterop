''' <summary>
''' Wrapper for COM root node, typically the application
''' </summary>
''' <typeparam name="TComObject">A type of the application (usually used with COM interop assemblies) or System.Object (usually used with COM objects created with CreateObject)</typeparam>
Public Class ComRootObject(Of TComObject)
    Inherits ComObjectBase

    ''' <summary>
    ''' Create a new COM root object e.g. for Excel.Application
    ''' </summary>
    ''' <param name="obj"></param>
    ''' <param name="onClosingAction">A close action, usually a method call to quit the application</param>
    Public Sub New(obj As TComObject, onClosingAction As OnClosingAction)
        Me.New(obj, Nothing, onClosingAction, Nothing)
    End Sub

    Public Sub New(obj As TComObject,
                           onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(Nothing, obj, onDisposeChildrenAction, onClosingAction, onClosedAction)
    End Sub

    Protected Friend Sub New(parentItemResponsibleForDisposal As Global.CompuMaster.ComInterop.ComObjectBase, obj As TComObject, onClosingAction As OnClosingAction)
        MyBase.New(parentItemResponsibleForDisposal, obj, Nothing, onClosingAction, Nothing)
    End Sub

    Protected Friend Sub New(parentItemResponsibleForDisposal As Global.CompuMaster.ComInterop.ComObjectBase, obj As TComObject,
                           onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
        MyBase.New(parentItemResponsibleForDisposal, obj, onDisposeChildrenAction, onClosingAction, onClosedAction)
    End Sub

    ''' <summary>
    ''' Close/quit the application
    ''' </summary>
    ''' <remarks>Identical as calling method Dispose() directly</remarks>
    Public Sub Close()
        Me.Dispose()
    End Sub

    ''' <summary>
    ''' The COM object with its accessible members
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ComObjectStronglyTyped As TComObject
        Get
            Return CType(MyBase.ComObject, TComObject)
        End Get
    End Property

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class
