Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Global.CompuMaster.MsPowerPointComInterop

    Friend NotInheritable Class PresentationWindow
        Private Const NativeObjectModel As UInteger = &HFFFFFFF0UI
        Private Shared ReadOnly DispatchId As New Guid("00020400-0000-0000-C000-000000000046")
        Private Delegate Function WindowCallback(window As IntPtr, data As IntPtr) As Boolean

        Friend Shared Function FindCandidates(matchesActivation As Func(Of IntPtr, Boolean)) As List(Of IntPtr)
            Dim candidates As New List(Of IntPtr)
            EnumWindows(
                Function(window, data)
                    Dim name As New StringBuilder(32)
                    GetClassName(window, name, name.Capacity)
                    If name.ToString().StartsWith("PPTFrameClass", StringComparison.Ordinal) AndAlso matchesActivation(window) Then
                        candidates.Add(window)
                    End If
                    Return True
                End Function, IntPtr.Zero)
            Return candidates
        End Function

        Friend Shared Function HasCaption(window As IntPtr, caption As String) As Boolean
            Dim title As New StringBuilder(1024)
            GetWindowText(window, title, title.Capacity)
            If title.ToString().Contains(caption) Then Return True

            ' The native PowerPoint document window class varies with the Office version.
            Dim matched = False
            EnumChildWindows(window,
                Function(child, data)
                    Dim name As New StringBuilder(32)
                    GetClassName(child, name, name.Capacity)
                    Select Case name.ToString()
                        Case "paneClassDC", "mdiClass"
                            matched = NativeApplicationHasCaption(child, caption)
                    End Select
                    Return Not matched
                End Function, IntPtr.Zero)
            Return matched
        End Function

        Private Shared Function NativeApplicationHasCaption(window As IntPtr, caption As String) As Boolean
            Dim dispatch = DispatchId
            Dim nativeWindow = IntPtr.Zero
            Dim documentWindow As Object = Nothing
            Dim application As Object = Nothing
            Try
                If AccessibleObjectFromWindow(window, NativeObjectModel, dispatch, nativeWindow) <> 0 OrElse nativeWindow = IntPtr.Zero Then Return False
                documentWindow = Marshal.GetObjectForIUnknown(nativeWindow)
                application = documentWindow.GetType().InvokeMember("Application", BindingFlags.GetProperty, Nothing, documentWindow, Nothing)
                If application Is Nothing Then Return False
                Dim value = TryCast(application.GetType().InvokeMember("Caption", BindingFlags.GetProperty, Nothing, application, Nothing), String)
                Return String.Equals(value, caption, StringComparison.Ordinal)
            Catch ex As Exception When TypeOf ex Is COMException OrElse TypeOf ex Is TargetInvocationException OrElse TypeOf ex Is MissingMemberException
                Return False
            Finally
                If application IsNot Nothing Then Marshal.ReleaseComObject(application)
                If documentWindow IsNot Nothing Then Marshal.ReleaseComObject(documentWindow)
                If nativeWindow <> IntPtr.Zero Then Marshal.Release(nativeWindow)
            End Try
        End Function

        <DllImport("user32.dll")>
        Private Shared Function EnumWindows(callback As WindowCallback, data As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll")>
        Private Shared Function EnumChildWindows(parent As IntPtr, callback As WindowCallback, data As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function GetClassName(window As IntPtr, name As StringBuilder, capacity As Integer) As Integer
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
        Private Shared Function GetWindowText(window As IntPtr, text As StringBuilder, capacity As Integer) As Integer
        End Function

        <DllImport("oleacc.dll")>
        Private Shared Function AccessibleObjectFromWindow(window As IntPtr, objectId As UInteger, ByRef interfaceId As Guid, ByRef value As IntPtr) As Integer
        End Function
    End Class
End Namespace
