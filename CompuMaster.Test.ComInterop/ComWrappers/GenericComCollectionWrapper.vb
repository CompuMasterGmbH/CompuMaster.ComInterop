﻿'Imports CompuMaster.ComInterop

'Public Class GenericComCollectionWrapper(Of TCollectionComObject As Class, TParentWrapper As ComObjectBase, TChildItemComObject As Class, TChildItemWrapper)
'    Inherits ObjectReadOnlyCollectionBase(Of PseudoComCollectionObject, PseudoComCollectionWrapper, PseudoComRootNodeWrapper, PseudoComItemObject, PseudoComItemWrapper)

'    Public Sub New(parentItemResponsibleForDisposal As PseudoComRootNodeWrapper, createdComObjectInstance As PseudoComCollectionObject)
'        MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
'    End Sub

'    Public Overrides ReadOnly Property Count As Integer
'        Get
'            Return MyBase.ComObjectStronglyTyped.Count
'        End Get
'    End Property

'    Public Overrides ReadOnly Property Item(index As Integer) As PseudoComItemWrapper
'        Get
'            Dim Wrapper As New PseudoComItemWrapper(Me, Me.ComObjectStronglyTyped.Item(index))
'            Return Wrapper
'        End Get
'    End Property

'    Protected Overrides Sub OnDisposeChildren()
'    End Sub

'    Protected Overrides Sub OnClosing()
'    End Sub

'    Protected Overrides Sub OnClosed()
'    End Sub

'End Class
