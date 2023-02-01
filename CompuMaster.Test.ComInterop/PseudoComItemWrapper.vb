Imports CompuMaster.ComInterop

Public Class PseudoComItemWrapper
    Inherits ComChildObject(Of PseudoComItemObject, PseudoComCollectionWrapper)

    Public Sub New(parentItemResponsibleForDisposal As PseudoComCollectionWrapper, createdComObjectInstance As PseudoComItemObject)
        MyBase.New(parentItemResponsibleForDisposal, createdComObjectInstance)
    End Sub

    Public ReadOnly Property Name As String
        Get
            Return Me.ComObjectStronglyTyped.Name
        End Get
    End Property

End Class
