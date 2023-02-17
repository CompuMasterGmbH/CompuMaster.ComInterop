Option Explicit On
Option Strict On

Namespace CompuMaster.ComInterop

    ''' <summary>
    ''' An application wrapper class (based on ComRootObject) for safe COM object handling and release for COM servers running on local machine
    ''' </summary>
    ''' <remarks>
    ''' PLEASE NOTE:
    ''' <list type="bullet">
    ''' <item>Since COM servers are tracked in process list of current machine, only COM servers running on local machine are supported. COM servers running on different machines are not supported.</item>
    ''' <item>After disposal, the finalizer waits until the process of the COM application disappears. In case of timeout, the process will be killed.</item>
    ''' </list>
    ''' </remarks>
    Public Class ComApplication(Of TComApplication)
        Inherits CompuMaster.ComInterop.ComRootObject(Of TComApplication)

        ''' <summary>
        ''' The process name which is used as filter before comparing process IDs to lookup the correct process
        ''' </summary>
        Private ReadOnly ExpectedProcessName As String

        ''' <summary>
        ''' The window handle (typically the hwnd property) of a COM server application which is used to identify the correct process on local machine
        ''' </summary>
        ''' <returns></returns>
        Public Delegate Function HwndOfComApplicationInstanceAction(comApplicationObject As ComApplication(Of TComApplication)) As Integer

        ''' <summary>
        ''' Run all code provided for OnClosing, e.g. a quit command
        ''' </summary>
        ''' <param name="comApplicationObject"></param>
        Public Delegate Sub OnApplicationClosingAction(comApplicationObject As ComApplication(Of TComApplication))

        ''' <summary>
        ''' Create a new MS Excel instance within its wrapper instance
        ''' </summary>
        ''' <param name="comApplicationInstance">A COM server object instance</param>
        ''' <param name="hwndOfComApplicationInstanceAction">The window handle (typically the hwnd property) of a COM server application which is used to identify the correct process on local machine</param>
        ''' <param name="expectedProcessName">The process name which is used as filter before comparing process IDs to lookup the correct process</param>
        Public Sub New(comApplicationInstance As TComApplication, hwndOfComApplicationInstanceAction As HwndOfComApplicationInstanceAction, expectedProcessName As String)
            Me.New(comApplicationInstance, hwndOfComApplicationInstanceAction, Nothing, expectedProcessName)
        End Sub

        ''' <summary>
        ''' Create a new MS Excel instance within its wrapper instance
        ''' </summary>
        ''' <param name="comApplicationInstance">A COM server object instance</param>
        ''' <param name="hwndOfComApplicationInstanceAction">The window handle (typically the hwnd property) of a COM server application which is used to identify the correct process on local machine</param>
        ''' <param name="expectedProcessName">The process name which is used as filter before comparing process IDs to lookup the correct process</param>
        Public Sub New(comApplicationInstance As TComApplication, hwndOfComApplicationInstanceAction As HwndOfComApplicationInstanceAction, onClosingAction As OnApplicationClosingAction, expectedProcessName As String)
            MyBase.New(comApplicationInstance, Sub(a) onClosingAction(CType(a, ComApplication(Of TComApplication))))
            Me.ExpectedProcessName = expectedProcessName
            Try
                Dim ExcelProcessID As Integer = Nothing
                ComTools.GetWindowThreadProcessId(hwndOfComApplicationInstanceAction(Me), ExcelProcessID)
                Me.ProcessId = ExcelProcessID
            Catch
            End Try
        End Sub

#If NETFRAMEWORK Or NETCOREAPP Then
        Public Sub New(comApplicationName As String, hwndOfComApplicationInstanceAction As HwndOfComApplicationInstanceAction, expectedProcessName As String)
            Me.New(comApplicationName, hwndOfComApplicationInstanceAction, Nothing, expectedProcessName)
        End Sub

        Public Sub New(comApplicationName As String, hwndOfComApplicationInstanceAction As HwndOfComApplicationInstanceAction, onClosingAction As OnApplicationClosingAction, expectedProcessName As String)
            MyBase.New(CType(ComTools.CreateComApplication(comApplicationName), TComApplication), Sub(a) onClosingAction(CType(a, ComApplication(Of TComApplication))))
            Me.ExpectedProcessName = expectedProcessName
            Try
                Dim ExcelProcessID As Integer = Nothing
                ComTools.GetWindowThreadProcessId(hwndOfComApplicationInstanceAction(Me), ExcelProcessID)
                Me.ProcessId = ExcelProcessID
            Catch
            End Try
        End Sub
#End If

        ''' <summary>
        ''' The process ID of the COM server
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property ProcessId As Integer

        ''' <summary>
        ''' The process of the COM server
        ''' </summary>
        ''' <returns></returns>
        Public Function Process() As System.Diagnostics.Process
            If Me.ProcessId = 0 Then
                Return Nothing
            Else
                Return System.Diagnostics.Process.GetProcessById(Me.ProcessId)
            End If
        End Function

        ''' <summary>
        ''' Required actions after the COM object has been closed, e.g. removing from a list of open documents
        ''' </summary>
        Protected Overrides Sub OnClosed()
            MyBase.OnClosed()
            CompuMaster.ComInterop.ComTools.ReleaseComObject(Me.ComObject)
            SafelyCloseExcelAppInstanceInternal()
        End Sub

        ''' <summary>
        ''' A timeout value for closing MS Excel regulary, default to 15 seconds
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>After timeout, process will be killed if the process hasn't exited</remarks>
        Public Property Timeout1AfterAppClosing As New TimeSpan(0, 0, 15)

        ''' <summary>
        ''' A timeout value for process exiting after MS Excel process has been killed, defaults to 1 second
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks>After timeout, code waits for disappearance of process in process list</remarks>
        Public Property Timeout2ProcessExitAfterAppKill As New TimeSpan(0, 0, 1)

        ''' <summary>
        ''' A timeout value for watching process list for disappeared MS Excel process, defaults to 1 second
        ''' </summary>
        ''' <returns>After timeout, process is expected to be closed "with chance of 99.99%" (not guaranteed)</returns>
        Public Property Timeout3ProcessListDisappearanceAfterAppKill As New TimeSpan(0, 0, 1)

        ''' <summary>
        ''' At some unkown circumstances, MS Excel process wasn't closed sometimes and required a forced process killing
        ''' </summary>
        Private Sub SafelyCloseExcelAppInstanceInternal()
            If ProcessId <> Nothing AndAlso Process() IsNot Nothing AndAlso Process.HasExited = False Then
                Try
                    ComTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, Timeout1AfterAppClosing) 'Sometimes it takes time to close MS Excel...
                    Me.Process.Refresh()
                    If Me.Process.HasExited = False Then
                        'Force kill on Excel 
                        Me.Process.Kill()
                        Try
                            ComTools.WaitUntilTrueOrTimeout(Function() Me.Process.HasExited = True, Timeout2ProcessExitAfterAppKill)
                        Catch 'ex As ArgumentException
                            'expected for invalid processId after kill
                        End Try
                        Try
                            ComTools.WaitUntilTrueOrTimeout(Function()
                                                                Dim ExcelProcesses() = System.Diagnostics.Process.GetProcessesByName(ExpectedProcessName)
                                                                For Each ExcelProcess In ExcelProcesses
                                                                    If ExcelProcess.Id = Me.ProcessId Then Return False
                                                                Next
                                                                Return True
                                                            End Function, Timeout3ProcessListDisappearanceAfterAppKill)
                        Catch 'ex As Exception
                            'ignore any exceptions on getting process list
                        End Try
                    End If
                Catch 'ex As Exception
                    'ignore any exceptions of watching/handling process close/kill
                End Try
            End If
        End Sub

        Public ReadOnly Property IsDisposed As Boolean
            Get
                Return MyBase.IsDisposedComObject
            End Get
        End Property

    End Class

End Namespace