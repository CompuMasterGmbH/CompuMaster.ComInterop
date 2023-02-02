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

        <SetUp>
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

        Private Shared Function IsPlatformSupportingComInterop() As Boolean
            Select Case System.Environment.OSVersion.Platform
                Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows
                    Return True
                Case Else
                    Return False
            End Select
        End Function

        Private Shared Function CreateExcelAppViaCom() As TestClassForExcelApp
            Dim Result As TestClassForExcelApp = Nothing
            If IsPlatformSupportingComInterop() Then
                'Expected to run successful
                Try
                    Result = New TestClassForExcelApp()
                Catch ex As Exception
                    Assert.Ignore("Windows platform with COM support, but COM application not available: " & ex.Message)
                End Try
                Assert.NotNull(Result)
                Result.InvokePropertySet("Visible", False)
                Result.InvokePropertySet("Interactive", False)
                Result.InvokePropertySet("ScreenUpdating", False)
                Result.InvokePropertySet("DisplayAlerts", False)
            Else
                Assert.Throws(Of Exception)(
                        Sub()
                            Result = New TestClassForExcelApp()
                        End Sub,
                        "CreateObject/COM not supported on non-windows platforms")
                Assert.Null(Result)
                Assert.Ignore("CreateObject/COM not supported on non-windows platforms")
            End If
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
            Tools.WaitUntilTrueOrTimeout(Function() Tools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 15))
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Excel-COM closed, but still " & Tools.ExcelProcesses.Length & " Excel processes running on this machine")

            TestClassForExcelApp.Close()
            GC.Collect(2, GCCollectionMode.Forced)
            GC.WaitForPendingFinalizers()
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Closing an Excel-COM for the 2nd time still doesn't change anything")

            Dim JitExcelApp As New Global.CompuMaster.ComInterop.ComRootObject(Of Object)(CreateObject("Excel.Application"), Sub(instance) instance.InvokeMethod("Quit"))
            Assert.AreEqual(1, Tools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
            JitExcelApp.Dispose()
            GC.Collect(2, GCCollectionMode.Forced)
            GC.WaitForPendingFinalizers()
            Tools.WaitUntilTrueOrTimeout(Function() Tools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 15))
            Assert.AreEqual(0, Tools.ExcelProcesses.Length, "Excel-COM closed, but still " & Tools.ExcelProcesses.Length & " Excel processes running on this machine")
        End Sub

        <Test>
        Public Sub CreateObjectFailing()
            Dim ComObj As Object
            Assert.Throws(Of Exception)(
                    Sub()
#Disable Warning CA1416 ' Diese Aufrufsite ist auf allen Plattformen erreichbar
                        ComObj = CreateObject("NeverGiveUp.ApplicationWontExist")
#Enable Warning CA1416 ' Diese Aufrufsite ist auf allen Plattformen erreichbar
                    End Sub,
                    "CreateObject/COM must fail on all platforms / COM application must not exist")
        End Sub

    End Class

End Namespace