Imports NUnit.Framework

<NonParallelizable>
Public MustInherit Class MsExcelTestBase
    Inherits ComTestBase

    <OneTimeSetUp>
    Public Sub OneTimeSetup()
        Console.WriteLine("OneTimeSetup: " & ExcelProcessTools.ExcelProcesses.Count & " Excel processes found")
        If ExcelProcessTools.ExcelProcesses.Length <> 0 Then
            Assert.Fail("Tests can't be executed while Excel processes are started on this machine")
        End If
        AssertIsComSupportedAndMsExcelAppInstalled()
    End Sub

    <SetUp>
    Public Sub Setup()
        Console.WriteLine("Setup: " & ExcelProcessTools.ExcelProcesses.Count & " Excel processes found")
        If ExcelProcessTools.ExcelProcesses.Length <> 0 Then
            Assert.Fail("Tests can't be executed while Excel processes are started on this machine")
        End If
    End Sub

    <TearDown>
    Public Sub TearDown()
        CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
        ExcelProcessTools.KillAllExcelProcesses() 'Kill all left-overs
    End Sub

    <Test>
    Public Sub NoExcelProcessesStarted()
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
    End Sub

    Protected Shared Sub AssertIsComSupportedAndMsExcelAppInstalled()
        If IsPlatformSupportingComInterop() = False Then
            Assert.Ignore("Platform not supported for COM")
        Else
            'Windows platform ok - MS Excel installed?
            Try
#Disable Warning CA1416
                Dim MsExcelType As Type = Type.GetTypeFromProgID("Excel.Application")
#Enable Warning CA1416
                Assert.IsNotNull(MsExcelType)
            Catch ex As Exception
                Assert.Ignore("MS Excel not installed: " & ex.Message)
            End Try
        End If
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Expected no MS Excel processes after checkup of installed MS Excel application")
    End Sub

End Class
