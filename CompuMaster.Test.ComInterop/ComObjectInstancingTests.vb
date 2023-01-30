Imports NUnit.Framework

Namespace CompuMaster.Test.ComInterop

    <NonParallelizable>
    Public Class ComObjectInstancingTests

        <OneTimeSetUp>
        Public Sub OneTimeSetup()
            Console.WriteLine("OneTimeSetup: " & Tools.ExcelProcesses.Count & " Excel processes found")
            If Tools.ExcelProcesses.Length <> 0 Then
                Assert.Fail("Tests can't be executed while Excel processes are started on this machine")
            End If
        End Sub

        <OneTimeSetUp>
        Public Sub Setup()
            Console.WriteLine("Setup: " & Tools.ExcelProcesses.Count & " Excel processes found")
            If Tools.ExcelProcesses.Length <> 0 Then
                Assert.Fail("Tests can't be executed while Excel processes are started on this machine")
            End If
        End Sub

        <TearDown>
        Public Sub TearDown()
            GC.Collect(2, GCCollectionMode.Forced)
            GC.WaitForPendingFinalizers()
            Tools.KillAllExcelProcesses() 'Kill all left-overs
        End Sub

        Private Shared Function CreateExcelAppViaCom() As TestClassForExcelApp
            Dim Result As TestClassForExcelApp = Nothing
            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows
                    'Expected to run successful
                    Result = New TestClassForExcelApp()
                Case Else
                    Assert.Throws(Of Exception)(
                        Sub()
                            Result = New TestClassForExcelApp()
                        End Sub,
                        "CreateObject/COM not supported on non-windows platforms")
            End Select
            Assert.NotNull(Result)
            Return Result
        End Function

        <Test>
        Public Sub NoExcelProcessesStarted()
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
        End Sub

        <Test>
        Public Sub ExcelProcessesStartedAndClosed()
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")

            Dim TestClassForExcelApp As TestClassForExcelApp = CreateExcelAppViaCom()
            Assert.AreEqual(1, Tools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")

            TestClassForExcelApp.Close()
            GC.Collect(2, GCCollectionMode.Forced)
            GC.WaitForPendingFinalizers()
            System.Threading.Thread.Sleep(3000)
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Excel-COM closed, but still " & Tools.ExcelProcesses.Length & " Excel processes running on this machine")

            TestClassForExcelApp.Close()
            GC.Collect(2, GCCollectionMode.Forced)
            GC.WaitForPendingFinalizers()
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Closing an Excel-COM for the 2nd time still doesn't change anything")
        End Sub

        <Test>
        Public Sub Dummy()
            Assert.Pass()
        End Sub

    End Class

End Namespace