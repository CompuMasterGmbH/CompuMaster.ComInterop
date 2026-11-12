Imports CompuMaster.ComInterop.Office
Imports System.Reflection
Imports System.Runtime.InteropServices

''' <summary>Wraps a Microsoft Word application without requiring the Office interop assembly.</summary>
Public Class WordApplication
    Inherits OfficeApplication

    ''' <summary>Activates Microsoft Word for COM automation.</summary>
    Public Sub New()
        MyBase.New("Word.Application", "WINWORD", False)
    End Sub

    ''' <inheritdoc/>
    Protected Overrides Function GetApplicationWindow() As IntPtr
        Dim window = InvokePropertyGet(Of Object)("ActiveWindow")
        If window Is Nothing Then Return IntPtr.Zero
        Try
            Return New IntPtr(CInt(window.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, Nothing, window, Nothing)))
        Finally
            Marshal.ReleaseComObject(window)
        End Try
    End Function

    ''' <inheritdoc/>
    Protected Overrides Sub QuitApplication()
        InvokeMethod("Quit", Type.Missing, Type.Missing, Type.Missing)
    End Sub
End Class
