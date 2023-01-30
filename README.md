# CompuMaster.ComInterop
COM interop library with safe design pattern for COM instancing

## When you need this COM interop library
1. COM interop often requires platform additional assemblies with platform dependency for Windows 32/64 bit
   * Microsoft Office development for Word/Excel/Powerpoint/etc. usualy requires you to add Microsoft's interop assemblies
   * Microsoft's interop assemblies often cause problems on client machines if there is another MS Office version installed than on the developer's machine
   * You just want to distribute your little small program to the clients regardless to their installed MS Office version
2. COM interop is often implented wrongly, causing MS Word/Excel/... to stay as process until your application terminates
   * Usually the application closes successfully if **all** COM objects have been closed and finalized correctly
   * Unfortunately, the .NET garbage collector can't do it automatically while your application is running
   * there is the need for finalizing all COM objects on dispose
3. After you created your COM instance e.g. with `CreateObject("Excel.Application")`, .NET isn't able to call many methods/properties/fields on COM objects using LateBinding, e.g. `MsExcelAppViaCom.Quit()`

## This nice assembly provides these features
* Access to public fields/properties/methods/function using System.Reflection (which allows calling of `MsExcelAppViaCom.Quit()` again)
* Provide a base class for your custom implementations to access all required fields/properties/methods/function from your application as simple as possible
* The base class comes with 
  * an implementation for the IDisposable interface to dispose/finalize all COM objects incl. their children COM objects correctly 
  * a design pattern which forces the developer to wrap all COM objects into classes inheriting from this base class
  
## Sample implementation for Microsoft Excel

See full sample at https://www.github.com/CompuMasterGmbH/CompuMaster.Excel/

<details>
<summary>First code impressions from https://www.github.com/CompuMasterGmbH/CompuMaster.Excel/</summary>

### ExcelApplication

```vb.net
Public Class ExcelApplication
    Inherits ComObjectBase

    Public Sub New()
        MyBase.New(Nothing, CreateObject("Excel.Application"))
        Me.Workbooks = New ExcelWorkbooksCollection(Me, Me)
    End Sub

    Public ReadOnly Property Workbooks As ExcelWorkbooksCollection

    Public Property UserControl As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("UserControl")
        End Get
        Set(value As Boolean)
            InvokePropertySet("UserControl", value)
        End Set
    End Property

    Public Property DisplayAlerts As Boolean
        Get
            Return InvokePropertyGet(Of Boolean)("DisplayAlerts")
        End Get
        Set(value As Boolean)
            InvokePropertySet("DisplayAlerts", value)
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

    Public Function Dialogs(type As Enumerations.XlBuiltInDialog) As ExcelDialog
        Return New ExcelDialog(Me, InvokePropertyGet("Dialogs", CType(type, Integer)))
    End Function

    Public Function Run(vbaMethodNameInclWorkbookName As String) As Object
        Return InvokeFunction("Run", New Object() {vbaMethodNameInclWorkbookName})
    End Function

    Public Function Run(workbookName As String, vbaMethod As String) As Object
        Return InvokeFunction("Run", New Object() {"'" & workbookName & "'!" & vbaMethod})
    End Function

    Public ReadOnly Property IsClosed As Boolean
        Get
            Return MyBase.IsDisposedComObject
        End Get
    End Property

    Public Sub Close()
        Me.Quit()
    End Sub

    Public Sub Quit()
        If Not IsDisposedComObject Then
            UserControl = True
            MyBase.CloseAndDisposeChildrenAndComObject()
        End If
    End Sub

    Private AdditionalDisposeChildrenList As New List(Of ComObjectBase)

    Protected Overrides Sub OnDisposeChildren()
        If Me.Workbooks IsNot Nothing Then Me.Workbooks.Dispose()
    End Sub

    Protected Overrides Sub OnClosing()
        InvokeMethod("Quit")
    End Sub

    Protected Overrides Sub OnClosed()
        GC.Collect(2, GCCollectionMode.Forced, True)
    End Sub

End Class
```

### Excel WorkboksCollection
```vb.net
Public Class ExcelWorkbooksCollection
    Inherits ComObjectBase

    Friend Sub New(parentItemResponsibleForDisposal As ComObjectBase, app As ExcelApplication)
        MyBase.New(parentItemResponsibleForDisposal, app.InvokePropertyGet("Workbooks"))
        Me.Parent = app
    End Sub

    Friend ReadOnly Parent As ExcelApplication

    Public Workbooks As New List(Of ExcelWorkbook)

    Public Function Open(path As String) As ExcelWorkbook
        Dim wb As New ExcelWorkbook(Me, Me, path)
        Me.Workbooks.Add(wb)
        Return wb
    End Function

    Protected Overrides Sub OnDisposeChildren()
        For MyCounter As Integer = Workbooks.Count - 1 To 0 Step -1
            Workbooks(MyCounter).Dispose()
        Next
    End Sub

    Protected Overrides Sub OnClosing()
    End Sub

    Protected Overrides Sub OnClosed()
    End Sub

End Class
```
</details>

## Your custom COM class

Usualy provides for your custom implementation
* Quit/Close application/document and dispose/finalize related COM objects
  * `base.CloseAndDisposeChildrenAndComObject()` 
* Invoke members of COM object by Reflection (instead of late binding) with several overloads of
  * `base.Invoke...`
* Direct access to the COM object
  * `base.ComObject`
* Status information on Com object if in-use vs. closed/disposed
  * `base.IsDisposedComObject`

Usualy requires you to implement
* constructor method (`Sub New` in VisualBasic)
* `OnDisposeChildren()`
  * Close and dispose commands for children objects
* `OnClosing()`
  * Required close commands for the COM object like App.Quit() or Document.Close()
* `OnClosed()`
  * Required actions after the COM object has been closed, e.g. removing from a list of open documents
