Namespace CompuMaster.ComInterop

    ''' <summary>
    ''' Safe design and implementation for all disposing and finalizing of COM objects and all their child objects
    ''' </summary>
    Public MustInherit Class ComObjectBase
        Implements IDisposable

        Protected Sub New(parentItemResponsibleForDisposal As ComObjectBase, createdComObjectInstance As Object,
                          onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction)
            If createdComObjectInstance Is Nothing Then Throw New ArgumentNullException(NameOf(createdComObjectInstance))
            _ComObject = createdComObjectInstance
            ComObjectType = _ComObject.GetType
            If parentItemResponsibleForDisposal IsNot Nothing AndAlso parentItemResponsibleForDisposal.RegisteredComChildren.Contains(Me) = False Then
                parentItemResponsibleForDisposal.RegisteredComChildren.Add(Me)
            End If
            Me._OnDisposeChildrenAction = onDisposeChildrenAction
            Me._OnClosingAction = onClosingAction
            Me._OnClosedAction = onClosedAction
        End Sub

        Private _ComObject As Object
        ''' <summary>
        ''' The COM object
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ComObject As Object
            Get
                Return _ComObject
            End Get
        End Property

        ''' <summary>
        ''' The ComObject's runtime-type
        ''' </summary>
        ''' <returns></returns>
        Protected ReadOnly Property ComObjectType As Type

        ''' <summary>
        ''' Close/quit the application inclusive its children COM objects
        ''' </summary>
        ''' <exception cref="Exception">If actions fail to close the COM object or its children, an exception is thrown</exception>
        ''' <remarks>Close/dispose actions occur only if not yet closed; identical as calling method Dispose() directly</remarks>
        Protected Sub CloseAndDisposeChildrenAndComObject()
            Me.Dispose(True)
        End Sub

        ''' <summary>
        ''' Has the COM object already been closed and disposed
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property IsDisposedComObject As Boolean
            Get
                Return _ComObject Is Nothing
            End Get
        End Property

#Region "Invoke methods"
        ''' <summary>
        ''' Invoke function member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="values">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        ''' <returns></returns>
        Public Function InvokeFunction(Of T)(name As String, ParamArray values As Object()) As T
            Return CompuMaster.Reflection.PublicInstanceMembers.InvokeFunction(Of T)(_ComObject, ComObjectType, name, values)
        End Function

        ''' <summary>
        ''' Invoke method member
        ''' </summary>
        ''' <param name="name"></param>
        ''' <param name="values">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        Public Sub InvokeMethod(name As String, ParamArray values As Object())
            CompuMaster.Reflection.PublicInstanceMembers.InvokeMethod(_ComObject, ComObjectType, name, values)
        End Sub

        ''' <summary>
        ''' Invoke property-get member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <returns></returns>
        Public Function InvokePropertyGet(Of T)(name As String) As T
            Return CompuMaster.Reflection.PublicInstanceMembers.InvokePropertyGet(Of T)(_ComObject, ComObjectType, name)
        End Function

        ''' <summary>
        ''' Invoke property-get member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="propertyArrayItem">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        ''' <returns></returns>
        Public Function InvokePropertyGet(Of T)(name As String, propertyArrayItem As Object) As T
            Return CompuMaster.Reflection.PublicInstanceMembers.InvokePropertyGet(Of T)(_ComObject, ComObjectType, name, propertyArrayItem)
        End Function

        ''' <summary>
        ''' Invoke property-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        Public Sub InvokePropertySet(Of T)(name As String, value As T)
            CompuMaster.Reflection.PublicInstanceMembers.InvokePropertySet(Of T)(_ComObject, ComObjectType, name, value)
        End Sub

        ''' <summary>
        ''' Invoke property-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="values">Arguments for the called member, remember to use System.Reflection.Missing.Value where required</param>
        Public Sub InvokePropertySet(Of T)(name As String, values As T())
            CompuMaster.Reflection.PublicInstanceMembers.InvokePropertySet(Of T)(_ComObject, ComObjectType, name, values)
        End Sub

        ''' <summary>
        ''' Invoke field-get member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <returns></returns>
        Public Function InvokeFieldGet(Of T)(name As String) As T
            Return CompuMaster.Reflection.PublicInstanceMembers.InvokeFieldGet(Of T)(_ComObject, ComObjectType, name)
        End Function

        ''' <summary>
        ''' Invoke field-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="value"></param>
        Public Sub InvokeFieldSet(Of T)(name As String, value As T)
            CompuMaster.Reflection.PublicInstanceMembers.InvokeFieldSet(Of T)(_ComObject, ComObjectType, name, value)
        End Sub

        ''' <summary>
        ''' Invoke field-set member
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="name"></param>
        ''' <param name="values"></param>
        Public Sub InvokeFieldSet(Of T)(name As String, values As T())
            CompuMaster.Reflection.PublicInstanceMembers.InvokeFieldSet(Of T)(_ComObject, ComObjectType, name, values)
        End Sub
#End Region

        ''' <summary>
        ''' Create a wrapper for a COM child object (e.g. a Workbooks collection) and register it for automatic disposal with this instance
        ''' </summary>
        ''' <typeparam name="TChildComObject"></typeparam>
        ''' <param name="comObject"></param>
        ''' <returns>The wrapper class of the COM child</returns>
        <Obsolete("Use CreateWrapperAndRegisterComChildForDisposal instead", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Function CreateWrapperAndRegisterComChildForDispoal(Of TChildComObject As Class)(comObject As TChildComObject) As ComChildObject(Of ComObjectBase, TChildComObject)
            Dim ChildWrapper As New ComChildObject(Of ComObjectBase, TChildComObject)(Me, comObject)
            Me.RegisteredComChildren.Add(ChildWrapper)
            Return ChildWrapper
        End Function

        ''' <summary>
        ''' Create a wrapper for a COM child object (e.g. a Workbooks collection) and register it for automatic disposal with this instance
        ''' </summary>
        ''' <typeparam name="TChildComObject"></typeparam>
        ''' <param name="comObject"></param>
        ''' <returns>The wrapper class of the COM child</returns>
        <Obsolete("Use CreateWrapperAndRegisterComChildForDisposal instead", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Function CreateWrapperAndRegisterComChildForDispoal(Of TChildComObject As Class)(comObject As TChildComObject, onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction) As ComChildObject(Of ComObjectBase, TChildComObject)
            Dim ChildWrapper As New ComChildObject(Of ComObjectBase, TChildComObject)(Me, comObject, onDisposeChildrenAction, onClosingAction, onClosedAction)
            Me.RegisteredComChildren.Add(ChildWrapper)
            Return ChildWrapper
        End Function

        ''' <summary>
        ''' Create a wrapper for a COM child object (e.g. a Workbooks collection) and register it for automatic disposal with this instance
        ''' </summary>
        ''' <typeparam name="TChildComObject"></typeparam>
        ''' <param name="comObject"></param>
        ''' <returns>The wrapper class of the COM child</returns>
        Public Function CreateWrapperAndRegisterComChildForDisposal(Of TChildComObject As Class)(comObject As TChildComObject) As ComChildObject(Of ComObjectBase, TChildComObject)
            Dim ChildWrapper As New ComChildObject(Of ComObjectBase, TChildComObject)(Me, comObject)
            Me.RegisteredComChildren.Add(ChildWrapper)
            Return ChildWrapper
        End Function

        ''' <summary>
        ''' Create a wrapper for a COM child object (e.g. a Workbooks collection) and register it for automatic disposal with this instance
        ''' </summary>
        ''' <typeparam name="TChildComObject"></typeparam>
        ''' <param name="comObject"></param>
        ''' <returns>The wrapper class of the COM child</returns>
        Public Function CreateWrapperAndRegisterComChildForDisposal(Of TChildComObject As Class)(comObject As TChildComObject, onDisposeChildrenAction As OnDisposeChildrenAction, onClosingAction As OnClosingAction, onClosedAction As OnClosedAction) As ComChildObject(Of ComObjectBase, TChildComObject)
            Dim ChildWrapper As New ComChildObject(Of ComObjectBase, TChildComObject)(Me, comObject, onDisposeChildrenAction, onClosingAction, onClosedAction)
            Me.RegisteredComChildren.Add(ChildWrapper)
            Return ChildWrapper
        End Function

        ''' <summary>
        ''' Register an independent wrapper class to be disposed when this object disposes
        ''' </summary>
        ''' <param name="childWrapper"></param>
        <Obsolete("Use RegisterComChildForDisposal instead", False)>
        <System.ComponentModel.EditorBrowsable(ComponentModel.EditorBrowsableState.Never)>
        Public Sub RegisterComChildForDispoal(childWrapper As ComObjectBase)
            Me.RegisteredComChildren.Add(childWrapper)
        End Sub

        ''' <summary>
        ''' Register an independent wrapper class to be disposed when this object disposes
        ''' </summary>
        ''' <param name="childWrapper"></param>
        Public Sub RegisterComChildForDisposal(childWrapper As ComObjectBase)
            Me.RegisteredComChildren.Add(childWrapper)
        End Sub

#Region "IDisposable"
        Private RegisteredComChildren As New List(Of ComObjectBase)

        ''' <summary>
        ''' Actions before close and dispose commands for children objects and this object 
        ''' </summary>
        Protected MustOverride Sub OnBeforeClosing()

        ''' <summary>
        ''' Close and dispose commands for children objects
        ''' </summary>
        Protected MustOverride Sub OnDisposeChildren()

        ''' <summary>
        ''' Required close commands for the COM object like App.Quit() or Document.Close()
        ''' </summary>
        Protected MustOverride Sub OnClosing()

        ''' <summary>
        ''' Required actions after the COM object has been closed, e.g. removing from a list of open documents
        ''' </summary>
        Protected MustOverride Sub OnClosed()

        Private _OnDisposeChildrenAction As OnDisposeChildrenAction
        ''' <summary>
        ''' Run all code provided for OnDisposeChildren
        ''' </summary>
        ''' <param name="instance"></param>
        Public Delegate Sub OnDisposeChildrenAction(instance As ComObjectBase)

        Private _OnClosingAction As OnClosingAction
        ''' <summary>
        ''' Run all code provided for OnClosing, e.g. a quit command
        ''' </summary>
        ''' <param name="instance"></param>
        Public Delegate Sub OnClosingAction(instance As ComObjectBase)

        Private _OnClosedAction As OnClosedAction
        ''' <summary>
        ''' Run all code provided for OnClosed, e.g. cleanup of collections or caches referencing this object
        ''' </summary>
        ''' <param name="instance"></param>
        Public Delegate Sub OnClosedAction(instance As ComObjectBase)

        ''' <summary>
        ''' Run all code provided for OnDisposeChildren
        ''' </summary>
        Private Sub _OnDisposeChildren()
            If _OnDisposeChildrenAction IsNot Nothing Then _OnDisposeChildrenAction(Me) 'run delegated method (if provided)
            OnDisposeChildren() 'run override-method
        End Sub

        ''' <summary>
        ''' Run all code provided for OnClosing, e.g. a quit command
        ''' </summary>
        Private Sub _OnClosing()
            If _OnClosingAction IsNot Nothing Then _OnClosingAction(Me) 'run delegated method (if provided)
            OnClosing() 'run override-method
        End Sub

        ''' <summary>
        ''' Run all code provided for OnClosed, e.g. cleanup of collections or caches referencing this object
        ''' </summary>
        Private Sub _OnClosed()
            If _OnClosedAction IsNot Nothing Then _OnClosedAction(Me) 'run delegated method (if provided)
            OnClosed() 'run override-method
        End Sub

        ''' <summary>
        ''' Run close and dispose for all children objects
        ''' </summary>
        Private Sub DisposeRegisteredComChildren()
            For MyCounter As Integer = 0 To Me.RegisteredComChildren.Count - 1
                Me.RegisteredComChildren(MyCounter).Dispose()
            Next
        End Sub

        Private disposedValue As Boolean
        Private isGC As Boolean

        ''' <summary>
        ''' Ignore exceptions caused by InvokeMethod calls to invalid objects (for safety and stability of application to not crash because of a failing finalizer)
        ''' </summary>
        Protected Property IgnoreMissingMethodExceptionsOnFinalize As Boolean = True

        ''' <summary>
        ''' Close and dispose the COM object and all of its children (if not yet done)
        ''' </summary>
        ''' <param name="disposing">True if called by method Dispose, False if called by method Finalize</param>
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If _ComObject IsNot Nothing Then
                    OnBeforeClosing()
                End If

                If disposing Then
                    If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                        Try
                            _OnDisposeChildren()
                            DisposeRegisteredComChildren()
                        Catch ex As Exception
                            'ignore
                            WriteToLogFileInUnitTestMode("ERROR on OnDisposeChildren: " & ex.ToString() & System.Environment.NewLine)
                        End Try
                    Else
                        _OnDisposeChildren()
                        DisposeRegisteredComChildren()
                    End If
                End If

                If _ComObject IsNot Nothing Then
                    'Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalizer überschreiben
                    If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                        Try
                            _OnClosing()
                        Catch ex As System.MissingMethodException
                            'ignore
                            WriteToLogFileInUnitTestMode("ERROR on OnClosing: " & ex.ToString() & System.Environment.NewLine)
                        Catch ex As Exception
                            'ignore
                            WriteToLogFileInUnitTestMode("ERROR on OnClosing: " & ex.ToString() & System.Environment.NewLine)
                        End Try
                    Else
                        _OnClosing()
                    End If
                    If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                        Try
                            ComTools.ReleaseComObject(_ComObject)
                        Catch ex As Exception
                            'ignore
                            WriteToLogFileInUnitTestMode("ERROR on ReleaseComObject: " & ex.ToString() & System.Environment.NewLine)
                        End Try
                    Else
                        ComTools.ReleaseComObject(_ComObject)
                    End If
                    If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                        Try
                            _OnClosed()
                        Catch ex As Exception
                            'ignore
                            WriteToLogFileInUnitTestMode("ERROR on OnClosed: " & ex.ToString() & System.Environment.NewLine)
                        End Try
                    Else
                        _OnClosed()
                    End If
                    'Große Felder auf NULL setzen
                    _ComObject = Nothing
                End If

                disposedValue = True
            End If
        End Sub

        ''' <summary>
        ''' Create a log entry in file GarbageCollectorLogOutput
        ''' </summary>
        ''' <param name="text"></param>
        Private Shared Sub WriteToLogFileInUnitTestMode(text As String)
            Try
                If GarbageCollectorLogOutput IsNot Nothing Then
                    System.IO.File.WriteAllText(GarbageCollectorLogOutput.FullName, text & System.Environment.NewLine)
                End If
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' Exceptions in dispose/finalize will be logged into this file
        ''' </summary>
        ''' <returns></returns>
        Friend Shared Property GarbageCollectorLogOutput As System.IO.FileInfo

        Protected Overrides Sub Finalize()
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
            isGC = True
            Dispose(disposing:=False)
            MyBase.Finalize()
        End Sub

        ''' <summary>
        ''' Garbage collector edition: close and dispose the COM object and all of its children (if not yet done)
        ''' </summary>
        ''' <remarks>If <see cref="IgnoreMissingMethodExceptionsOnFinalize"/> is true, suppress exceptions if closing actions of this object or children objects fail</remarks>
        Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
            isGC = True
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

        ''' <summary>
        ''' Close/quit the application inclusive its children COM objects
        ''' </summary>
        ''' <exception cref="Exception">If actions fail to close the COM object or its children, an exception is thrown</exception>
        ''' <remarks>Close/dispose actions occur only if not yet closed</remarks>
        Public Sub Dispose()
            ' Ändern Sie diesen Code nicht. Fügen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
            Dispose(disposing:=True)
        End Sub
#End Region

    End Class

End Namespace