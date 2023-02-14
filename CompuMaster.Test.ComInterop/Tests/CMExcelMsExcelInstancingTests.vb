Imports NUnit.Framework
Imports CompuMaster.Excel.MsExcelCom

<NonParallelizable>
Public Class CMExcelMsExcelInstancingTests
    Inherits MsExcelTestBase

    <Test>
    Public Sub CreateMsExcelInstanceWithExplicitDispose()
        Dim Excel As New MsExcelApplicationWrapper()
        Assert.NotNull(Excel)
        Excel.Dispose()
    End Sub

    <Test>
    Public Sub CreateMsExcelInstanceWithImplicitDispose()
        Dim Excel As New MsExcelApplicationWrapper()
        Assert.NotNull(Excel)
    End Sub


    <Test>
    Public Sub CreateMsExcelInstanceAndCreateNewWorkbook()
        Dim Excel As New MsExcelApplicationWrapper()
        Dim Wb = Excel.Workbooks.Add
        Assert.IsNotEmpty(Wb.ComObjectStronglyTyped.Name)
    End Sub


End Class
