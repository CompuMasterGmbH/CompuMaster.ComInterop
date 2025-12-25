Option Explicit On
Option Strict On

Namespace CompuMaster.ComInterop

    ''' <summary>
    ''' COM interop tools
    ''' </summary>
    Public NotInheritable Class ComTools

        ''' <summary>
        ''' Release a COM object (using System.Runtime.InteropServices.Marshal.ReleaseComObject respectively FinalReleaseComObject)
        ''' </summary>
        ''' <param name="obj"></param>
        Public Shared Sub ReleaseComObject(ByVal obj As Object)
            Try
                If obj IsNot Nothing Then
                    Dim RemainingComReferencesToRelease As Integer = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                    Do While RemainingComReferencesToRelease > 0
                        RemainingComReferencesToRelease = System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                    Loop
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj)
                    obj = Nothing
                End If
            Catch
                obj = Nothing
            End Try
        End Sub

        ''' <summary>
        ''' Run a forced garbage collection and wait for pending finalizers
        ''' </summary>
        Public Shared Sub GarbageCollectAndWaitForPendingFinalizers()
            GC.Collect(0, GCCollectionMode.Forced, True)
            GC.Collect(2, GCCollectionMode.Forced, True)
            GC.WaitForPendingFinalizers()
        End Sub

#If NETCOREAPP Or NETFRAMEWORK Then
        ''' <summary>
        ''' Create a new instance of a COM server
        ''' </summary>
        ''' <param name="name"></param>
        ''' <returns></returns>
        ''' <exception cref="PlatformNotSupportedException">Platform windows is known to support COM, platforms Linux/Unix or Mac are known to not support COM, or instancing failed for other reasons</exception>
        ''' <exception cref="ComApplicationNotAvailableException">Instancing failed with a COMException, typically interpreted as</exception>
        Public Shared Function CreateComApplication(name As String, Optional serverName As String = "") As Object
            Try
#Disable Warning CA1416
                Dim ComAppType As Type = Type.GetTypeFromProgID(name)
#Enable Warning
                If ComAppType Is Nothing Then
                    Throw New ComApplicationNotAvailableException("COM application not available: " & name, Nothing)
                Else
                    Return Interaction.CreateObject(name)
                End If
            Catch ex As PlatformNotSupportedException
                Throw
            Catch ex As System.Runtime.InteropServices.COMException
                Throw New ComApplicationNotAvailableException(ex.Message, ex)
            Catch ex As Exception
                Throw New PlatformNotSupportedException(ex.Message, ex)
            End Try
        End Function
#End If

        ''' <summary>
        ''' Is the current platform known to support COM servers (windows) and is the COM application available in current environment
        ''' </summary>
        ''' <param name="name">Name of a registered COM server object</param>
        ''' <returns></returns>
        Public Shared Function IsPlatformSupportingComInteropAndMsExcelAppInstalled(name As String) As Boolean
            If IsPlatformSupportingComInterop() = False Then
                Return False
            Else
                'Windows platform ok - application installed?
                Try
#Disable Warning CA1416
                    Dim ComAppType As Type = Type.GetTypeFromProgID(name)
#Enable Warning
                    Return ComAppType IsNot Nothing
                Catch
                    Return False
                End Try
            End If
        End Function

        ''' <summary>
        ''' Is the current platform known to support COM servers (windows) or to not support COM servers (e.g. Linux/Unix/Mac)
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function IsPlatformSupportingComInterop() As Boolean
            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' A function or lambda expression for testing a result
        ''' </summary>
        ''' <returns></returns>
        Public Delegate Function WaitUntilTrueConditionTest() As Boolean

        ''' <summary>
        ''' Wait until expression is true or timeout
        ''' </summary>
        ''' <param name="expression"></param>
        ''' <param name="maxTimeout"></param>
        ''' <returns>True if expression was true, False if timeout</returns>
        Public Shared Function WaitUntilTrueOrTimeout(expression As WaitUntilTrueConditionTest, maxTimeout As TimeSpan) As Boolean
            Dim Start As DateTime = DateTime.Now
            Do While DateTime.Now.Subtract(Start) < maxTimeout
                If expression() = True Then
                    Return True
                End If
                If maxTimeout.TotalDays > 0 OrElse maxTimeout.Hours > 0 Then 'prevent exceeded range when calling maxTimeout.TotalMilliseconds
                    'Check at least twice per second
                    System.Threading.Thread.Sleep(500)
                Else
                    'Check at least 10 times and minimum twice per second
                    System.Threading.Thread.Sleep(System.Math.Min(CType(maxTimeout.TotalMilliseconds / 10, Integer), 500))
                End If
            Loop
            Return False
        End Function

        Friend Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger

    End Class

End Namespace