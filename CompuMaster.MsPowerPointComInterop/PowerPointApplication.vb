Imports CompuMaster.ComInterop.Office

''' <summary>Wraps a Microsoft PowerPoint application without requiring the Office interop assembly.</summary>
Public Class PowerPointApplication
    Inherits OfficeApplication

    ''' <summary>Activates Microsoft PowerPoint for COM automation.</summary>
    Public Sub New()
        MyBase.New("PowerPoint.Application", "POWERPNT", False)
    End Sub

    ''' <inheritdoc/>
    Protected Overrides Function GetApplicationWindow() As IntPtr
        If HasExistingApplication Then Return IntPtr.Zero
        Dim candidates = PresentationWindow.FindCandidates(AddressOf WindowMatchesActivation)
        If candidates.Count = 0 Then Return IntPtr.Zero

        ' A temporary unique caption associates the COM object with its exact window.
        Dim original = InvokePropertyGet(Of String)("Caption")
        Dim marker = Guid.NewGuid().ToString("N")
        Try
            InvokePropertySet("Caption", marker)
            Return candidates.Find(Function(window) PresentationWindow.HasCaption(window, marker))
        Finally
            If String.Equals(InvokePropertyGet(Of String)("Caption"), marker, StringComparison.Ordinal) Then
                InvokePropertySet("Caption", original)
            End If
        End Try
    End Function
End Class
