Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports CompuMaster.ComInterop.Office
Imports NUnit.Framework

Public Class OfficeDiagnosticLogsTest

    <TestCase("WINWORD")>
    <TestCase("POWERPNT")>
    <TestCase("MSACCESS")>
    Public Sub RejectsAnotherOfficeProductEvenWhenThePidMatches(processName As String)
        If Environment.OSVersion.Platform <> PlatformID.Win32NT Then Assert.Ignore("Windows file identities required.")
        Dim root = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(root)
        Try
            Using current = Process.GetCurrentProcess()
                Dim correctFile = WritePrimary(root, processName, current.Id)
                WritePrimary(root, "EXCEL", current.Id)
                Dim mixedFile = WritePrimary(root, processName, current.Id)
                File.AppendAllText(mixedFile, Record("EXCEL", current.Id))
                Dim tracker As New OfficeDiagnosticLogs(current, root, processName)

                Assert.That(tracker.Snapshot().Select(Function(file) file.FullName), [Is].EquivalentTo({correctFile}))
            End Using
        Finally
            Directory.Delete(root, recursive:=True)
        End Try
    End Sub

    Private Shared Function WritePrimary(root As String, processName As String, pid As Integer) As String
        Dim logPath = Path.Combine(root, "Primary1_" & Guid.NewGuid().ToString("D") & ".log")
        File.WriteAllText(logPath, Record(processName, pid))
        Return logPath
    End Function

    Private Shared Function Record(processName As String, pid As Integer) As String
        Return DateTime.UtcNow.ToString("MM/dd/yyyy HH:mm:ss.fff", CultureInfo.InvariantCulture) & vbTab &
            processName & " (0x" & pid.ToString("X", CultureInfo.InvariantCulture) & ")" & vbTab & "message" & vbCrLf
    End Function
End Class
