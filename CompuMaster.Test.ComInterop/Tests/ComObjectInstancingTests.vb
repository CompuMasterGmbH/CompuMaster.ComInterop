Imports NUnit.Framework

<NonParallelizable>
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
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")

        Dim TestClassForExcelApp As TestClassForExcelApp = CreateExcelAppViaCom()
        Assert.AreEqual(1, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")

        TestClassForExcelApp.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 15))
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Excel-COM closed, but still " & ExcelProcessTools.ExcelProcesses.Length & " Excel processes running on this machine")

        TestClassForExcelApp.Close()
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Closing an Excel-COM for the 2nd time still doesn't change anything")

#Disable Warning CA1416
        Dim JitExcelApp As New Global.CompuMaster.ComInterop.ComRootObject(Of Object)(CreateObject("Excel.Application"), Sub(instance) instance.InvokeMethod("Quit"))
#Enable Warning CA1416
        Assert.AreEqual(1, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
        JitExcelApp.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 15))
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Excel-COM closed, but still " & ExcelProcessTools.ExcelProcesses.Length & " Excel processes running on this machine")
    End Sub

    <Test>
    Public Sub ExcelProcessesConcurrentlyStartedAndClosed()
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")

        Dim TestClassForExcelApp1 As TestClassForExcelApp = CreateExcelAppViaCom()
        Assert.AreEqual(1, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
        Assert.AreEqual(False, TestClassForExcelApp1.Visible)

        Dim TestClassForExcelApp2 As TestClassForExcelApp = CreateExcelAppViaCom()
        Assert.AreEqual(2, ExcelProcessTools.ExcelProcesses.Length, "Tests can't be executed while Excel processes are started on this machine")
        Assert.AreEqual(False, TestClassForExcelApp2.Visible)

        TestClassForExcelApp2.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ExcelProcesses.Length = 1, New TimeSpan(0, 0, 15))
        Assert.AreEqual(1, ExcelProcessTools.ExcelProcesses.Length, "Excel-COM closed, but still " & ExcelProcessTools.ExcelProcesses.Length & " Excel processes running on this machine")
        Assert.Catch(Of System.Exception)(Function()
                                              Return TestClassForExcelApp2.Visible
                                          End Function, "Property shouldn't be accessible any more")
        Assert.AreEqual(False, TestClassForExcelApp1.Visible, "Property should be accessible further more")

        TestClassForExcelApp1.Close()
        ExcelProcessTools.WaitUntilTrueOrTimeout(Function() ExcelProcessTools.ExcelProcesses.Length = 0, New TimeSpan(0, 0, 15))
        Assert.AreEqual(0, ExcelProcessTools.ExcelProcesses.Length, "Excel-COM closed, but still " & ExcelProcessTools.ExcelProcesses.Length & " Excel processes running on this machine")
        Assert.Catch(Of System.Exception)(Function()
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