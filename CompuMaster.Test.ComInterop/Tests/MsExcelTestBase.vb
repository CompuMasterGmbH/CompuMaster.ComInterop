Imports NUnit.Framework

<Parallelizable(ParallelScope.All)>
Public MustInherit Class MsExcelTestBase
    Inherits ComTestBase

    Public ReadOnly AutoRunGarbageCollector As Boolean = False

    Public Sub New()
        MyBase.New
    End Sub

    Public Sub New(autoRunGarbageCollector As Boolean)
        Me.New
        Me.AutoRunGarbageCollector = autoRunGarbageCollector
    End Sub

    <OneTimeSetUp>
    Public Sub OneTimeSetup()
        AssertIsComSupportedAndMsExcelAppInstalled()
    End Sub

    <SetUp>
    Public Sub Setup()
    End Sub

    <TearDown>
    Public Sub TearDown()
    End Sub

    <OneTimeTearDown>
    Public Sub OneTimeTearDown()
        If Me.AutoRunGarbageCollector Then 'Should not be required any more - especially in parallelly running tests
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
        End If
    End Sub

    Protected Shared Sub AssertIsComSupportedAndMsExcelAppInstalled()
        If IsPlatformSupportingComInterop() = False Then
            Assert.Ignore("Platform not supported for COM")
        Else
            'Windows platform ok - MS Excel installed?
            Try
#Disable Warning CA1416
                Dim MsExcelType As Type = Type.GetTypeFromProgID("Excel.Application")
#Enable Warning CA1416
                Assert.IsNotNull(MsExcelType)
            Catch ex As Exception
                Assert.Ignore("MS Excel not installed: " & ex.Message)
            End Try
        End If
    End Sub

End Class
