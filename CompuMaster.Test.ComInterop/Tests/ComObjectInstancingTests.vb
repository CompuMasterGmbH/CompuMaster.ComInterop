Imports System.Text
Imports CompuMaster.ComInterop
Imports NUnit.Framework

<Parallelizable(ParallelScope.All)>
Public Class ComObjectInstancingTests
    Inherits MsExcelTestBase

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
    Public Sub ExcelProcessesStartedAndClosed()
        Dim TestClassForExcelApp As TestClassForExcelApp = CreateExcelAppViaCom()
        ExcelProcessTools.AssertValidOpenedProcess(TestClassForExcelApp.ProcessID)

        TestClassForExcelApp.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ProcessIdExists(TestClassForExcelApp.ProcessID) = False, New TimeSpan(0, 0, 15))
        ExcelProcessTools.AssertValidClosedProcess(TestClassForExcelApp.ProcessID)

        TestClassForExcelApp.Close()
        ExcelProcessTools.AssertValidClosedProcess(TestClassForExcelApp.ProcessID)

#Disable Warning CA1416
        Dim JitExcelApp As New Global.CompuMaster.ComInterop.ComRootObject(Of Object)(CreateObject("Excel.Application"), Sub(instance) instance.InvokeMethod("Quit"))
#Enable Warning CA1416
        Dim JitExcelAppExcelAppHwnd = JitExcelApp.InvokePropertyGet(Of Integer)("Hwnd")
        Dim JitExcelAppProcessID = ComTools.LookupProcessID(JitExcelAppExcelAppHwnd)
        ExcelProcessTools.AssertValidOpenedProcess(JitExcelAppProcessID)

        JitExcelApp.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ProcessIdExists(JitExcelAppProcessID) = False, New TimeSpan(0, 0, 15))
        ExcelProcessTools.AssertValidClosedProcess(JitExcelAppProcessID)
    End Sub

    Public Shared Function ExcelProcessesConcurrentlyStartedAndClosed_Runs() As IEnumerable
        Return Enumerable.Range(1, 20)
    End Function

    <Test>
    <TestCaseSource(NameOf(ExcelProcessesConcurrentlyStartedAndClosed_Runs))>
    Public Sub ExcelProcessesConcurrentlyStartedAndClosed(runId As Integer)
        TestContext.WriteLine($"Run {runId}, Thread {Threading.Thread.CurrentThread.ManagedThreadId}")

        Dim TestClassForExcelApp1 As TestClassForExcelApp = CreateExcelAppViaCom()
        ExcelProcessTools.AssertValidOpenedProcess(TestClassForExcelApp1.ProcessID)
        Assert.AreEqual(False, TestClassForExcelApp1.Visible)

        Dim TestClassForExcelApp2 As TestClassForExcelApp = CreateExcelAppViaCom()
        ExcelProcessTools.AssertValidOpenedProcess(TestClassForExcelApp2.ProcessID)
        Assert.AreEqual(False, TestClassForExcelApp2.Visible)

        TestClassForExcelApp2.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ProcessIdExists(TestClassForExcelApp2.ProcessID) = False, New TimeSpan(0, 0, 60))
        ExcelProcessTools.AssertValidOpenedProcess(TestClassForExcelApp1.ProcessID)
        ExcelProcessTools.AssertValidClosedProcess(TestClassForExcelApp2.ProcessID)
        Assert.Throws(Of CompuMaster.Reflection.InvokeException)(Function()
                                                                     Return TestClassForExcelApp2.Visible
                                                                 End Function, "Property shouldn't be accessible any more")
        Assert.AreEqual(False, TestClassForExcelApp1.Visible, "Property should be accessible further more")

        TestClassForExcelApp1.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ProcessIdExists(TestClassForExcelApp1.ProcessID) = False, New TimeSpan(0, 0, 60))
        ExcelProcessTools.AssertValidClosedProcess(TestClassForExcelApp1.ProcessID)
        Assert.Throws(Of CompuMaster.Reflection.InvokeException)(Function()
                                                                     Return TestClassForExcelApp1.Visible
                                                                 End Function, "Property shouldn't be accessible any more")
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