Imports NUnit.Framework

Public Class KillExcelProcesses

    <Test, Explicit("Run on demand, only")>
    Public Sub ForceKillForAllExcelProcesses()
        If ExcelProcessTools.ExcelProcesses.Length = 0 Then Assert.Ignore("Test can be executed ONLY while Excel processes are started on this machine")
        ExcelProcessTools.KillAllExcelProcesses(True)
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "All excel processes have been killed, but still they are available; tests can't be executed while Excel processes are started on this machine")
    End Sub

End Class