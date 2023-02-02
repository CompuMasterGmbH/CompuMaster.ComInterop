Public MustInherit Class ComObjectBase
    Implements IDisposable

    Protected Sub New(parentItemResponsibleForDisposal As ComObjectBase, createdComObjectInstance As Object)
        If createdComObjectInstance Is Nothing Then Throw New ArgumentNullException(NameOf(createdComObjectInstance))
        _ComObject = createdComObjectInstance
        ComObjectType = _ComObject.GetType
        If parentItemResponsibleForDisposal IsNot Nothing AndAlso parentItemResponsibleForDisposal.RegisteredComChildren.Contains(Me) = False Then
            parentItemResponsibleForDisposal.RegisteredComChildren.Add(Me)
        End If
    End Sub

    Private _ComObject As Object
    Public ReadOnly Property ComObject As Object
        Get
            Return _ComObject
        End Get
    End Property

    Protected ReadOnly Property ComObjectType As Type

    ''' <summary>
    ''' Close children and COM object (without suppression of exceptions)
    ''' </summary>
    ''' <remarks>Executes only if not yet closed</remarks>
    Protected Sub CloseAndDisposeChildrenAndComObject()
        If Not Me.IsDisposedComObject Then
            Me.Dispose(True)
        End If
    End Sub

    Protected ReadOnly Property IsDisposedComObject As Boolean
        Get
            Return _ComObject Is Nothing
        End Get
    End Property

    'Public Function InvokeFunction(name As String, ParamArray values As Object()) As Object
    '    Return ReflectionTools.InvokeFunction(_ComObject, ComObjectType, name, values)
    'End Function

    Public Function InvokeFunction(Of T)(name As String, ParamArray values As Object()) As T
        Return ReflectionTools.InvokeFunction(Of T)(_ComObject, ComObjectType, name, values)
    End Function

    Public Sub InvokeMethod(name As String, ParamArray values As Object())
        ReflectionTools.InvokeMethod(_ComObject, ComObjectType, name, values)
    End Sub

    'Public Function InvokePropertyGet(name As String) As Object
    '    Return ReflectionTools.InvokePropertyGet(_ComObject, ComObjectType, name)
    'End Function

    Public Function InvokePropertyGet(name As String, propertyArrayItem As Object) As Object
        Return ReflectionTools.InvokePropertyGet(_ComObject, ComObjectType, name, propertyArrayItem)
    End Function

    Public Function InvokePropertyGet(Of T)(name As String) As T
        Return ReflectionTools.InvokePropertyGet(Of T)(_ComObject, ComObjectType, name)
    End Function

    Public Function InvokePropertyGet(Of T)(name As String, propertyArrayItem As Object) As T
        Return ReflectionTools.InvokePropertyGet(Of T)(_ComObject, ComObjectType, name, propertyArrayItem)
    End Function

    Public Sub InvokePropertySet(Of T)(name As String, value As T)
        ReflectionTools.InvokePropertySet(Of T)(_ComObject, ComObjectType, name, value)
    End Sub

    Public Sub InvokePropertySet(Of T)(name As String, values As T())
        ReflectionTools.InvokePropertySet(Of T)(_ComObject, ComObjectType, name, values)
    End Sub

    'Public Function InvokeFieldGet(name As String) As Object
    '    Return ReflectionTools.InvokeFieldGet(_ComObject, ComObjectType, name)
    'End Function

    Public Function InvokeFieldGet(Of T)(name As String) As T
        Return ReflectionTools.InvokeFieldGet(Of T)(_ComObject, ComObjectType, name)
    End Function

    Public Sub InvokeFieldSet(Of T)(name As String, value As T)
        ReflectionTools.InvokeFieldSet(Of T)(_ComObject, ComObjectType, name, value)
    End Sub

    Public Sub InvokeFieldSet(Of T)(name As String, values As T())
        ReflectionTools.InvokeFieldSet(Of T)(_ComObject, ComObjectType, name, values)
    End Sub

    ''' <summary>
    ''' Create a wrapper for a COM child object (e.g. a Workbooks collection) and register it for automatic disposal with this instance
    ''' </summary>
    ''' <typeparam name="TChildComObject"></typeparam>
    ''' <param name="comObject"></param>
    ''' <returns>The wrapper class of the COM child</returns>
    Public Function CreateWrapperAndRegisterComChildForDispoal(Of TChildComObject As Class)(comObject As TChildComObject) As ComChildObject(Of TChildComObject, ComObjectBase)
        Dim ChildWrapper As New ComChildObject(Of TChildComObject, ComObjectBase)(Me, comObject)
        Me.RegisteredComChildren.Add(ChildWrapper)
        Return ChildWrapper
    End Function

    ''' <summary>
    ''' Register an independent wrapper class to be disposed when this object disposes
    ''' </summary>
    ''' <param name="childWrapper"></param>
    Public Sub RegisterComChildForDispoal(childWrapper As ComObjectBase)
        Me.RegisteredComChildren.Add(childWrapper)
    End Sub

#Region "IDisposable"
    Private RegisteredComChildren As New List(Of ComObjectBase)

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

    Private Sub DisposeRegisteredComChildren()
        For MyCounter As Integer = 0 To Me.RegisteredComChildren.Count - 1
            Me.RegisteredComChildren(MyCounter).Dispose()
        Next
    End Sub

    Private disposedValue As Boolean
    Private isGC As Boolean = False

    ''' <summary>
    ''' Ignore exceptions caused by InvokeMethod calls to invalid objects (for safety and stability of application to not crash because of a failing finalizer)
    ''' </summary>
    Protected IgnoreMissingMethodExceptionsOnFinalize As Boolean = True

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                    Try
                        OnDisposeChildren()
                        DisposeRegisteredComChildren()
                    Catch
                        'ignore
                    End Try
                Else
                    OnDisposeChildren()
                    DisposeRegisteredComChildren()
                End If
            End If

            If _ComObject IsNot Nothing Then
                'Nicht verwaltete Ressourcen (nicht verwaltete Objekte) freigeben und Finalizer �berschreiben
                If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                    Try
                        OnClosing()
                    Catch ex As System.MissingMethodException
                        'ignore
                    Catch
                        'ignore
                    End Try
                Else
                    OnClosing()
                End If
                ComTools.ReleaseComObject(_ComObject)
                If isGC And IgnoreMissingMethodExceptionsOnFinalize Then
                    Try
                        OnClosed()
                    Catch
                        'ignore
                    End Try
                Else
                    OnClosed()
                End If
                'Gro�e Felder auf NULL setzen
                _ComObject = Nothing
            End If

            disposedValue = True
        End If
    End Sub

    Protected Overrides Sub Finalize()
        ' �ndern Sie diesen Code nicht. F�gen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
        isGC = True
        Dispose(disposing:=False)
        MyBase.Finalize()
    End Sub

    Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
        ' �ndern Sie diesen Code nicht. F�gen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
        isGC = True
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub

    Public Sub Dispose()
        ' �ndern Sie diesen Code nicht. F�gen Sie Bereinigungscode in der Methode "Dispose(disposing As Boolean)" ein.
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
