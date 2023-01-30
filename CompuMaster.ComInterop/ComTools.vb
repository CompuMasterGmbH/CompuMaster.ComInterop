Option Explicit On
Option Strict On

''' <summary>
''' COM interop tools
''' </summary>
Public NotInheritable Class ComTools

    Public Shared Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch
            obj = Nothing
        End Try
    End Sub

End Class
