Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices

Namespace Global.CompuMaster.ComInterop.Office

    ''' <summary>Wraps an Office automation application and tracks its diagnostic logs.</summary>
    Public MustInherit Class OfficeApplication
        Inherits ComRootObject(Of Object)

        Private ReadOnly ProcessName As String
        Private ReadOnly StartedActivating As DateTime
        Private ReadOnly FinishedActivating As DateTime
        Private ReadOnly QuitWithoutProcess As Boolean
        Private Diagnostics As OfficeDiagnosticLogs
        Private PreserveLogs As Boolean?
        Private Finished As Boolean

        ''' <summary>Activates the Office application and enables diagnostic tracking.</summary>
        ''' <param name="progId">The COM programmatic identifier.</param>
        ''' <param name="processName">The expected Office executable name without extension.</param>
        ''' <param name="quitWithoutProcess">Whether this wrapper already owns its COM server independently of diagnostic tracking.</param>
        Protected Sub New(progId As String, processName As String, quitWithoutProcess As Boolean)
            Me.New(Activate(progId, processName), processName, quitWithoutProcess)
        End Sub

        Private Sub New(activation As Activation, processName As String, quitWithoutProcess As Boolean)
            MyBase.New(activation.Instance, Nothing)
            Me.ProcessName = processName
            Me.QuitWithoutProcess = quitWithoutProcess
            StartedActivating = activation.Started
            FinishedActivating = activation.Finished
            HasExistingApplication = activation.HasExistingApplication
        End Sub

        ''' <summary>Gets or sets whether diagnostic files are preserved byte for byte.</summary>
        Public Property PreserveDiagnosticLogFiles As Boolean
            Get
                Dim tracker = GetDiagnostics()
                Return If(tracker Is Nothing, PreserveLogs.GetValueOrDefault(), tracker.Preserve)
            End Get
            Set(value As Boolean)
                PreserveLogs = value
                Dim tracker = GetDiagnostics()
                If tracker IsNot Nothing Then tracker.Preserve = value
            End Set
        End Property

        ''' <summary>Returns existing diagnostic files associated with the automation process.</summary>
        Public ReadOnly Property DiagnosticLogFiles As ReadOnlyCollection(Of FileInfo)
            Get
                Dim tracker = GetDiagnostics()
                Return If(tracker Is Nothing, Array.AsReadOnly(Array.Empty(Of FileInfo)()), tracker.Snapshot())
            End Get
        End Property

        ''' <summary>Gets the window handle verified to belong to this COM application.</summary>
        ''' <returns>The window handle, or zero if no matching window is available.</returns>
        Protected MustOverride Function GetApplicationWindow() As IntPtr

        ''' <summary>Gets whether a process for this Office application was already running before activation.</summary>
        Protected ReadOnly Property HasExistingApplication As Boolean

        ''' <summary>Checks whether a window belongs to a process started during this COM activation.</summary>
        ''' <param name="window">A window handle to check.</param>
        ''' <returns>Whether the process matches the expected executable and activation interval.</returns>
        Protected Function WindowMatchesActivation(window As IntPtr) As Boolean
            Using process = FindActivatedProcess(window)
                Return process IsNot Nothing
            End Using
        End Function

        Private Function GetDiagnostics() As OfficeDiagnosticLogs
            If Diagnostics IsNot Nothing OrElse Finished OrElse IsClosed Then Return Diagnostics
            Try
                Using process = FindActivatedProcess(GetApplicationWindow())
                    If process Is Nothing Then Return Nothing
                    Diagnostics = OfficeDiagnosticLogs.ForProcess(process, ProcessName)
                    If Diagnostics IsNot Nothing AndAlso PreserveLogs.HasValue Then Diagnostics.Preserve = PreserveLogs.Value
                End Using
            Catch
                ' A missing window or inaccessible process prevents diagnostic tracking.
            End Try
            Return Diagnostics
        End Function

        Private Function FindActivatedProcess(window As IntPtr) As Process
            Dim pid = ComTools.LookupProcessID(window)
            If pid = 0 Then Return Nothing
            Dim process As Process = Nothing
            Try
                process = Process.GetProcessById(pid)
                Dim started = process.StartTime.ToUniversalTime()
                If started >= StartedActivating AndAlso started <= FinishedActivating AndAlso
                    String.Equals(process.ProcessName, ProcessName, StringComparison.OrdinalIgnoreCase) Then
                    Dim match = process
                    process = Nothing
                    Return match
                End If
            Catch
                ' The window or process may have disappeared.
            Finally
                If process IsNot Nothing Then process.Dispose()
            End Try
            Return Nothing
        End Function

        ''' <inheritdoc/>
        Protected Overrides Sub OnClosing()
            Dim tracker = Diagnostics
            If tracker Is Nothing AndAlso Not QuitWithoutProcess Then Return
            If tracker IsNot Nothing Then tracker.BeforeQuit()
            QuitApplication()
            If tracker IsNot Nothing Then tracker.AfterQuit()
        End Sub

        ''' <summary>Issues the Office application's quit command.</summary>
        Protected Overridable Sub QuitApplication()
            InvokeMethod("Quit")
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub ReleaseOwnedComReference(instance As Object)
            If QuitWithoutProcess Then
                MyBase.ReleaseOwnedComReference(instance)
            Else
                Marshal.ReleaseComObject(instance)
            End If
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Sub Dispose(disposing As Boolean)
            If Finished Then Return
            Dim tracker = GetDiagnostics()
            Dim succeeded = False
            Try
                MyBase.Dispose(disposing)
                Finished = True
                succeeded = True
            Finally
                If tracker IsNot Nothing Then tracker.Finish(succeeded, waitForExit:=disposing)
            End Try
        End Sub

        Private Shared Function Activate(progId As String, processName As String) As Activation
            If Not RuntimeInformation.IsOSPlatform(OSPlatform.Windows) Then Throw New PlatformNotSupportedException("Office COM automation requires Windows.")
            Dim hasExisting = False
            For Each existing In Process.GetProcessesByName(processName)
                Using existing
                    hasExisting = True
                End Using
            Next
            Dim started = DateTime.UtcNow
            Dim comType = Type.GetTypeFromProgID(progId, True)
            Dim instance = Activator.CreateInstance(comType)
            Return New Activation With {.Instance = instance, .Started = started, .Finished = DateTime.UtcNow, .HasExistingApplication = hasExisting}
        End Function

        Private Structure Activation
            Friend Instance As Object
            Friend Started As DateTime
            Friend Finished As DateTime
            Friend HasExistingApplication As Boolean
        End Structure
    End Class
End Namespace
