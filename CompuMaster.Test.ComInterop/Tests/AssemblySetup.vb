Imports NUnit.Framework

<SetUpFixture>
Public Class AssemblySetup

    <OneTimeSetUp>
    Public Sub GlobalSetup()
        ' Läuft GENAU EINMAL vor ALLEN Tests der Assembly
        TestContext.Progress.WriteLine("=== GLOBAL SETUP START ===")
        CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput = New System.IO.FileInfo(System.IO.Path.Combine(System.IO.Path.GetTempPath, "ComInterop.GarbageCollector.log"))
        If CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.Exists = False Then
            CompuMaster.ComInterop.ComObjectBase.GarbageCollectorLogOutput.Create.Close()
        End If
        TestContext.Progress.WriteLine("=== GLOBAL SETUP END ===")
    End Sub

    <OneTimeTearDown>
    Public Sub GlobalTeardown()
        ' Läuft GENAU EINMAL nach ALLEN Tests der Assembly
        'TestContext.Progress.WriteLine("=== GLOBAL TEARDOWN END ===")
    End Sub

End Class
