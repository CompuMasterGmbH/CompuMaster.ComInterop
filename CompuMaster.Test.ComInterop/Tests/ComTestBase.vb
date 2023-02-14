Public MustInherit Class ComTestBase

    Protected Shared Function IsPlatformSupportingComInterop() As Boolean
        Select Case System.Environment.OSVersion.Platform
            Case PlatformID.Win32NT, PlatformID.Win32S, PlatformID.Win32Windows
                Return True
            Case Else
                Return False
        End Select
    End Function

End Class
