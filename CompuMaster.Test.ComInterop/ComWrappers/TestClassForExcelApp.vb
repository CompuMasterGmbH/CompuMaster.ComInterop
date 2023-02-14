Public Class TestClassForExcelApp
    Inherits Global.CompuMaster.ComInterop.ComObjectBase

#Disable Warning CA1416 ' Diese Aufrufsite ist auf allen Plattformen erreichbar
    Public Sub New()
        MyBase.New(Nothing,
                   CreateObject("Excel.Application"),
                   Nothing,
                   Nothing,
                   Nothing)
    End Sub
#Enable Warning CA1416 ' Diese Aufrufsite ist auf allen Plattformen erreichbar

    Public Property UserControl As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("UserControl")
        End Get
        Set(value As Boolean)
            InvokePropertySet("UserControl", value)
        End Set
    End Property

    Public Property Visible As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("Visible")
        End Get
        Set(value As Boolean)
            InvokePropertySet("Visible", value)
        End Set
    End Property

    Public ReadOnly Property IsClosed As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Public Sub Close()
        MyBase.CloseAndDisposeChildrenAndComObject()
    End Sub

    Protected Overrides Sub OnDisposeChildren()
    End Sub

    Protected Overrides Sub OnClosing()
        InvokeMethod("Quit")
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

    Protected Overrides Sub OnBeforeClosing()
        UserControl = True
    End Sub

End Class