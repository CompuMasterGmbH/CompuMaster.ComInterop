Imports NUnit.Framework

Namespace CompuMaster.Test.ComInterop

    Public Class KillExcelProcesses

        <Test, Explicit("Run on demand, only")>
        Public Sub ForceKillForAllExcelProcesses()
            If Tools.ExcelProcesses.Length = 0 Then Assert.Ignore("Test can be executed ONLY while Excel processes are started on this machine")
            Tools.KillAllExcelProcesses(True)
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "All excel processes have been killed, but still they are available; tests can't be executed while Excel processes are started on this machine")
        End Sub

    End Class

End Namespace