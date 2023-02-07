Namespace CompuMaster.ComInterop

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
        ''' Has the COM object already been closed and disposed
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsClosed As Boolean
            Get
                Return MyBase.IsDisposedComObject
            End Get
        End Property

        ''' <summary>
        ''' Close/quit the application inclusive its children COM objects
        ''' </summary>
        ''' <exception cref="Exception">If actions fail to close the COM object or its children, an exception is thrown</exception>
        ''' <remarks>Close/dispose actions occur only if not yet closed; identical as calling method Dispose() directly</remarks>
        Public Sub Close()
            MyBase.CloseAndDisposeChildrenAndComObject()
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

        ''' <summary>
        ''' Close and dispose commands for children objects
        ''' </summary>
        Protected Overrides Sub OnDisposeChildren()
        End Sub

        ''' <summary>
        ''' Required close commands for the COM object like App.Quit() or Document.Close()
        ''' </summary>
        Protected Overrides Sub OnClosing()
        End Sub

        ''' <summary>
        ''' Required actions after the COM object has been closed, e.g. removing from a list of open documents
        ''' </summary>
        Protected Overrides Sub OnClosed()
        End Sub

        ''' <summary>
        ''' Actions before close and dispose commands for children objects and this object 
        ''' </summary>
        Protected Overrides Sub OnBeforeClosing()
        End Sub

    End Class

End Namespace