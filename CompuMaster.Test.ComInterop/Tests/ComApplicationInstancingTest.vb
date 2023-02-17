Imports NUnit.Framework

Public Class ComApplicationInstancingTest
    Inherits ComTestBase

    Class ExcelAppByName
        Inherits CompuMaster.ComInterop.ComApplication(Of Object)

        Public Sub New()
            MyBase.New("Excel.Application",
                       Function(x) x.InvokePropertyGet(Of Integer)("Hwnd"),
                       Sub(x) x.InvokeMethod("Quit"),
                       "EXCEL")
        End Sub

    End Class

    Class ExcelAppByType
        Inherits CompuMaster.ComInterop.ComApplication(Of Microsoft.Office.Interop.Excel.Application)

        Public Sub New()
            MyBase.New(New Microsoft.Office.Interop.Excel.Application,
                       Function(x) x.ComObjectStronglyTyped.Hwnd,
                       Sub(x) x.ComObjectStronglyTyped.Quit(),
                       "EXCEL")
        End Sub

    End Class

    <OneTimeSetUp>
    Public Sub OneTimeSetup()
        If MyBase.IsPlatformSupportingComInteropAndMsExcelAppInstalled("Excel.Application") = False Then
            Assert.Ignore("Platform not supported")
        End If
    End Sub

    <TearDown>
    Public Sub TearDown()
        CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
    End Sub

    <Test> Public Sub TestExcelAppByName()
        Dim E As New ExcelAppByName
        Assert.AreEqual("Microsoft Excel", E.InvokePropertyGet(Of String)("Name"))
        Assert.NotZero(E.ProcessId)
        E.Dispose()
    End Sub

    <Test> Public Sub TestExcelAppByInteropClass()
        Dim E As New ExcelAppByType
        Assert.AreEqual("Microsoft Excel", E.ComObjectStronglyTyped.Name)
        Assert.NotZero(E.ProcessId)
        E.Dispose()
    End Sub

    <Test> Public Sub TestJitExcelAppByName()
        Dim E As New CompuMaster.ComInterop.ComApplicationBase(Of Object)(
            "Excel.Application",
            Function(x) x.InvokePropertyGet(Of Integer)("Hwnd"),
            Sub(x) x.InvokeMethod("Quit"),
            "EXCEL")
        Assert.AreEqual("Microsoft Excel", E.InvokePropertyGet(Of String)("Name"))
        Assert.NotZero(E.ProcessId)
        E.Dispose()
    End Sub

    <Test> Public Sub TestJitExcelAppByInteropClass()
        Dim E As New CompuMaster.ComInterop.ComApplication(Of Microsoft.Office.Interop.Excel.Application)(
            New Microsoft.Office.Interop.Excel.Application,
            Function(x) x.ComObjectStronglyTyped.Hwnd,
            Sub(x) x.ComObjectStronglyTyped.Quit(),
            "EXCEL")
        Assert.AreEqual("Microsoft Excel", E.ComObjectStronglyTyped.Name)
        Assert.NotZero(E.ProcessId)
        E.Dispose()
    End Sub

End Class
