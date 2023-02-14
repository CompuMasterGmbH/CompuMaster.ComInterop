'Imports CompuMaster.ComInterop

'Public Class ObjectCollection(Of TChildComObject As Class, TParentWrapper As ComObjectBase, TChildWrapper As ComChildObject(Of TChildComObject, TParentWrapper))
'    Inherits ObjectCollectionBase(Of TChildComObject, TParentWrapper, TChildWrapper)

'    Public Sub New(parentItemResponsibleForDisposal As TParentWrapper, createdComObjectInstance As TChildComObject)
'        MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
'    End Sub

'    Private ItemsList As New List(Of TChildWrapper)

'    Public Overrides ReadOnly Property Count As Integer
'        Get
'            Return ItemsList.Count
'        End Get
'    End Property

'    Public Overrides ReadOnly Property Item(index As Integer) As TChildWrapper
'        Get
'            Return ItemsList.Item(index)
'        End Get
'    End Property

'    Protected Overrides Sub OnDisposeChildren()
'        For MyCounter As Integer = Count - 1 To 0 Step -1
'            ItemsList(MyCounter).Dispose()
'        Next
'    End Sub

'    Protected Overrides Sub OnClosing()
'    End Sub

'    Protected Overrides Sub OnClosed()
'    End Sub

'    Public Overrides Sub Add(item As TChildWrapper)
'        ItemsList.Add(item)
'    End Sub

'    Public Overrides Sub Insert(index As Integer, item As TChildWrapper)
'        ItemsList.Insert(index, item)
'    End Sub

'    Public Overrides Sub Remove(item As TChildWrapper)
'        ItemsList.Remove(item)
'    End Sub

'    Public Overrides Sub RemoveAt(index As Integer)
'        ItemsList.RemoveAt(index)
'    End Sub

'End Class
