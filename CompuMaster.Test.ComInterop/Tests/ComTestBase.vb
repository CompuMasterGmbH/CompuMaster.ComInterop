Imports NUnit.Framework

Public MustInherit Class ComTestBase

    Public Sub New()
        CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput = New System.IO.FileInfo(System.IO.Path.Combine(System.IO.Path.GetTempPath, "ComInterop.GarbageCollector.log"))
        If CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.Exists = False Then
            CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.Create.Close()
        End If
    End Sub

    Protected Shared Function IsPlatformSupportingComInterop() As Boolean
        Return CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInterop
    End Function

    Protected Shared Function IsPlatformSupportingComInteropAndMsExcelAppInstalled(name As String) As Boolean
        Return CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled(name)
    End Function

    <SetUp>
    Public Sub Setup_LogReview()
        System.Console.WriteLine("Garbage collector log at: " & CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.FullName)
    End Sub

    <OneTimeTearDown>
    Public Sub OneTimeTearDown_LogReview()
        If CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.Length <> 0 Then
            Throw New Exception("CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput file size <> 0, errors have been logged and must be reviewed")
        End If
    End Sub

    <Test>
    <Explicit("Only explicit file reset available/required")>
    Public Sub ResetCompuMasterComInteropComObjectBaseGarbageCollectorLogOutputFile()
        CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.Delete()
    End Sub

End Class
