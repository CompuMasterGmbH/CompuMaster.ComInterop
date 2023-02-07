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

    End Class

End Namespace