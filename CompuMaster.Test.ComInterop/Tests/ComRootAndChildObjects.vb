Imports NUnit.Framework
Imports CompuMaster.ComInterop

<Parallelizable>
Public Class ComRootAndChildObjects

    <Test>
    Public Sub TestComRoot()
        Dim ComRoot As New ComRootObject(Of String)("", Nothing)
        Assert.AreEqual("", ComRoot.ComObject)

        Dim ComChild As New ComChildObject(Of ComRootObject(Of String), String)(ComRoot, "")
        Assert.AreEqual("", ComChild.ComObject)
        Assert.IsNotNull(ComChild.Parent)

        'Dim ComChildCollection As New GenericComCollectionWrapper() '(Of String, ComRootObject(Of String), ComChildObject(Of PseudoComCollectionObject, ComRootObject(Of String)))(ComRoot, New PseudoComCollectionObject(3))
        'Assert.AreEqual("Manual Item 1", ComChildCollection.Item(0).ComObjectStronglyTyped.Item(0).Name)
        'Assert.AreEqual(3, ComChildCollection.Count)
        'ComChildCollection.Add(New ComChildObject(Of PseudoComItemObject, ComRootObject(Of String))(ComChildCollection, New PseudoComItemObject("Manual Item 1")))
        'Assert.AreEqual(4, ComChildCollection.Count)
        '
        'Dim ComChildItem1 As ObjectCollection(Of ComRootAndChildObjects, PseudoComItemObject) = ComChildCollection.Item(3)
        'Assert.AreEqual("Manual Item 1", ComChildItem1.ComObject.Name)

    End Sub

    <Test>
    Public Sub TestComRootStronglyTyped()
        Dim ComRoot As New PseudoComRootNodeWrapper(New PseudoComRootNodeObject())
        Assert.IsNotNull(ComRoot.ComObject)
        Assert.IsNotNull(ComRoot.ComObjectStronglyTyped)
        Assert.IsNotNull(ComRoot.Modules)
        Assert.IsNotNull(ComRoot.Modules.ComObject)
        Assert.IsNotNull(ComRoot.Modules.ComObjectStronglyTyped)
        Assert.AreEqual(ComRoot, ComRoot.Modules.Parent)
        Assert.AreEqual(3, ComRoot.Modules.Count)
        Assert.IsNotNull(ComRoot.Modules.Item(0))
        Assert.IsNotNull(ComRoot.Modules.Item(0).ComObject)
        Assert.IsNotNull(ComRoot.Modules.Item(0).ComObjectStronglyTyped)
        Assert.AreEqual(ComRoot.Modules, ComRoot.Modules.Item(0).Parent)
        Assert.AreEqual("List item 1", ComRoot.Modules.Item(0).Name)
        Assert.AreEqual("List item 2", ComRoot.Modules.Item(1).Name)
        Assert.AreEqual("List item 3", ComRoot.Modules.Item(2).Name)


        'Dim ComChild As New ComChildObject(Of String, ComRootObject(Of String), String)(ComRoot, "")
        'Assert.AreEqual("", ComChild.ComObject)
        'Assert.IsNotNull(ComChild.ComObjectStronglyTyped)
        'Assert.IsNotNull(ComChild.Parent)
        '
        'Dim ComChildCollection As New PseudoObjectCollection(Of String, ComRootObject, ComChildObject(Of String, ComRootObject(Of String)))(ComRoot, New Object())
        'Assert.Zero(ComChildCollection.Count)
        'ComChildCollection.Add(New ComChildObject(Of PseudoComItemObject, ComRootObject(Of String))(ComChildCollection, New PseudoComItemObject("Item1")))
        'Assert.Zero(ComChildCollection.Count)
        '
        'Dim ComChildItem1 As ObjectCollection(Of ComRootAndChildObjects, PseudoComItemObject) = ComChildCollection.Item(0)
        'Assert.AreEqual("Item1", ComChildItem1.ComObject.Name)
    End Sub

End Class
