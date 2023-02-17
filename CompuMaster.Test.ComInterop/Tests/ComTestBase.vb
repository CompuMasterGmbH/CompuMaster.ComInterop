Public MustInherit Class ComTestBase

    Protected Shared Function IsPlatformSupportingComInterop() As Boolean
        Return CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInterop
    End Function

    Protected Shared Function IsPlatformSupportingComInteropAndMsExcelAppInstalled(name As String) As Boolean
        Return CompuMaster.ComInterop.ComTools.IsPlatformSupportingComInteropAndMsExcelAppInstalled(name)
    End Function

End Class
