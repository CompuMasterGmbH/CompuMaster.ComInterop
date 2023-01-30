Public Class Tools

    Public Shared Function ExcelProcesses() As Process()
        Return System.Diagnostics.Process.GetProcessesByName("Excel")
    End Function

    Public Shared Sub KillAllExcelProcesses()
        GC.Collect(2, GCCollectionMode.Forced)
        GC.WaitForPendingFinalizers()
        Dim FoundExcelProcesses As Process() = Tools.ExcelProcesses
        For Each p In FoundExcelProcesses
            p.Close()
        Next
        If FoundExcelProcesses.Length <> 0 Then
            'Wait for Excel to close
            System.Threading.Thread.Sleep(3000)
        End If
    End Sub

    Public Delegate Function WaitUntilTrueConditionTest() As Boolean

    Public Shared Function WaitUntilTrueOrTimeout(test As WaitUntilTrueConditionTest, maxTimeout As TimeSpan) As Boolean
        Dim Start As DateTime = Now
        Do
            If test() = True Then
                Return True
            End If
            System.Threading.Thread.Sleep(500)
        Loop
        Return False
    End Function

End Class
