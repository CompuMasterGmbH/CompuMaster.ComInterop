Imports NUnit.Framework

Public Class ExcelProcessTools

    Public Shared Sub AssertValidProcessId(processID As Integer)
        If processID = 0 Then
            Assert.NotZero(processID, "Process ID <> 0 expected")
        End If
    End Sub

    Public Shared Sub AssertValidOpenedProcess(processID As Integer)
        AssertValidProcessId(processID)
        Dim FoundProcess = LookupProcess(processID)
        Assert.NotNull(FoundProcess, "Process with ID " & processID & " should exist")
        System.Console.WriteLine("Process ID " & processID & " " & If(FoundProcess IsNot Nothing, "found", "not found (any more)"))
    End Sub

    Public Shared Sub AssertValidClosedProcess(processID As Integer)
        AssertValidProcessId(processID)
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ProcessIdExists(processID) = False, New TimeSpan(0, 0, 60))
        Dim FoundProcess = LookupProcess(processID)
        Assert.Null(FoundProcess, "No process with ID " & processID & " should exist any more")
        System.Console.WriteLine("Process ID " & processID & " " & If(FoundProcess IsNot Nothing, "found", "not found (any more)"))
    End Sub

    Public Shared Function ProcessIdExists(processID As Integer) As Boolean
        Return LookupProcess(processID) IsNot Nothing
    End Function

    Public Shared Function LookupProcess(processID As Integer) As System.Diagnostics.Process
        If processID = 0 Then Throw New ArgumentNullException(NameOf(processID))
        Dim FoundProcess As System.Diagnostics.Process
        Try
            FoundProcess = System.Diagnostics.Process.GetProcessById(processID)
        Catch ex As System.ArgumentException
            'if process doesn't exist, expected System.ArgumentException : Es wird kein Prozess mit der ID 52336 ausgeführt.
            FoundProcess = Nothing
        End Try
        Return FoundProcess
    End Function

    <Obsolete("SHOULD NOT BE REQUIRED ANY MORE", False)>
    Public Shared Function ExcelProcesses() As Process()
        Return System.Diagnostics.Process.GetProcessesByName("Excel")
    End Function

    <Obsolete("SHOULD NOT BE REQUIRED ANY MORE", False)>
    Public Shared Sub KillAllExcelProcesses(Optional forceKill As Boolean = False)
        'Close by gargabe collector
        CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()

        'Wait for Excel to close - max 3 seconds
        WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 3))

        'Close processes
        Dim FoundExcelProcesses As Process() = ExcelProcessTools.ExcelProcesses
        For Each p In FoundExcelProcesses
            p.Close()
        Next

        If forceKill Then
            'Kill processes
            FoundExcelProcesses = ExcelProcessTools.ExcelProcesses
            For Each p In FoundExcelProcesses
                p.Kill()
            Next
        End If

        'Wait for Excel to close - max 5 seconds
        WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 5))
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
