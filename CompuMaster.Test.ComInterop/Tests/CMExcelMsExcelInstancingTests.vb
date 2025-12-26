Imports NUnit.Framework
Imports CompuMaster.Excel.MsExcelCom

<Parallelizable(ParallelScope.All)>
Public Class CMExcelMsExcelInstancingTests
    Inherits MsExcelTestBase

    <Test>
    Public Sub CreateMsExcelInstanceWithExplicitDispose()
        Dim ExcelProcessID As Integer = CreateMsExcelInstance(True)
        ExcelProcessTools.AssertValidClosedProcess(ExcelProcessID)
    End Sub

    <Test>
    Public Sub CreateMsExcelInstanceWithImplicitDispose()
        Dim ExcelProcessID As Integer = CreateMsExcelInstance(False)
        CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
        ExcelProcessTools.AssertValidClosedProcess(ExcelProcessID)
    End Sub

    Private Function CreateMsExcelInstance(explicitDispose As Boolean) As Integer
        Dim Excel As New MsExcelApplicationWrapper()
        System.Console.WriteLine(Excel.ToString)
        Assert.NotNull(Excel)
        Dim ExcelProcessID As Integer = Excel.ExcelProcessId
        ExcelProcessTools.AssertValidOpenedProcess(ExcelProcessID)
        If explicitDispose Then Excel.Dispose()
        Return ExcelProcessID
    End Function

    <Test>
    Public Sub CreateMsExcelInstanceAndCreateNewWorkbookWithExplicitDispose()
        Dim ExcelProcessID As Integer = CreateMsExcelInstanceAndCreateNewWorkbookWithImplicitDispose(True)
        ExcelProcessTools.AssertValidClosedProcess(ExcelProcessID)
    End Sub

    <Test>
    Public Sub CreateMsExcelInstanceAndCreateNewWorkbookWithImplicitDispose()
        Dim ExcelProcessID As Integer = CreateMsExcelInstanceAndCreateNewWorkbookWithImplicitDispose(False)
        CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
        ExcelProcessTools.AssertValidClosedProcess(ExcelProcessID)
        'ExcelProcessTools.KillAllExcelProcesses() 'Kill all left-overs
    End Sub

    Private Function CreateMsExcelInstanceAndCreateNewWorkbookWithImplicitDispose(explicitDispose As Boolean) As Integer
        Dim Excel As New MsExcelApplicationWrapper()
        Dim ProcessID As Integer = Excel.ExcelProcessId
        System.Console.WriteLine(Excel.ToString)
        ExcelProcessTools.AssertValidOpenedProcess(ProcessID)
        Dim Wb = Excel.Workbooks.Add
        Assert.IsNotEmpty(Wb.ComObjectStronglyTyped.Name)
        If explicitDispose Then Excel.Dispose()
        Return ProcessID
    End Function

    <Test>
    Public Sub CreateMsExcelInstanceAndCreateNewWorkbookWithImplicitDispose2()
        Dim Excel As New MsExcelApplicationWrapper()
        Dim ProcessID As Integer = Excel.ExcelProcessId
        System.Console.WriteLine(Excel.ToString)
        ExcelProcessTools.AssertValidOpenedProcess(ProcessID)
        Dim Wb = Excel.Workbooks.Add
        Assert.IsNotEmpty(Wb.ComObjectStronglyTyped.Name)
        If False Then Excel.Dispose()
    End Sub

End Class
