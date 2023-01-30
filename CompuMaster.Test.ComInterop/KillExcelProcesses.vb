Imports NUnit.Framework

Namespace CompuMaster.Test.ComInterop

    Public Class KillExcelProcesses

        <Test, Explicit("Run on demand, only")>
        Public Sub ForceKillForAllExcelProcesses()
            Assert.AreNotEqual(0, Tools.ExcelProcesses.Length, "Test can be executed ONLY while Excel processes are started on this machine")
            Tools.KillAllExcelProcesses()
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
        End Sub

    End Class

End Namespace