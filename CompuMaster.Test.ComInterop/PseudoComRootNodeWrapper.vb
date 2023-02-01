Imports CompuMaster.ComInterop

Public Class PseudoComRootNodeWrapper
    Inherits ComRootObject(Of PseudoComRootNodeObject)

    Public Sub New(comObject As PseudoComRootNodeObject)
        MyBase.New(comObject)
        Modules = New PseudoComCollectionWrapper(Me, New PseudoComCollectionObject(3))
    End Sub

    Public ReadOnly Property Modules As PseudoComCollectionWrapper

    Public Sub Quit()
        Me.ComObjectStronglyTyped.Quit()
    End Sub

End Class
