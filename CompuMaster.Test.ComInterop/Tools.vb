Public Class Tools

    Public Shared Function ExcelProcesses() As Process()
        Return System.Diagnostics.Process.GetProcessesByName("Excel")
    End Function

    Public Shared Sub KillAllExcelProcesses(Optional forceKill As Boolean = False)
        'Close by gargabe collector
        GC.Collect(2, GCCollectionMode.Forced)
        GC.WaitForPendingFinalizers()

        'Wait for Excel to close - max 3 seconds
        WaitUntilTrueOrTimeout(Function() Tools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 3))

        'Close processes
        Dim FoundExcelProcesses As Process() = Tools.ExcelProcesses
        For Each p In FoundExcelProcesses
            p.Close()
        Next

        If forceKill Then
            'Kill processes
            FoundExcelProcesses = Tools.ExcelProcesses
            For Each p In FoundExcelProcesses
                p.Kill()
            Next
        End If

        'Wait for Excel to close - max 5 seconds
        WaitUntilTrueOrTimeout(Function() Tools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 5))
    End Sub

    Public Delegate Function WaitUntilTrueConditionTest() As Boolean

    Public Shared Function WaitUntilTrueOrTimeout(test As WaitUntilTrueConditionTest, maxTimeout As TimeSpan) As Boolean
        Dim Start As DateTime = Now
        Do While Now.Subtract(Start) < maxTimeout
            If test() = True Then
                Return True
            End If
            System.Threading.Thread.Sleep(500)
        Loop
        Return False
    End Function

End Class
