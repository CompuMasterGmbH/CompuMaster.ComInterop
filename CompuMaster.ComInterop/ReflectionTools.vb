Option Explicit On
Option Strict On

''' <summary>
''' Provides simplified InvokeMember for public members of an instance object
''' </summary>
Public NotInheritable Class ReflectionTools

    Public Shared Function InvokeFunction(obj As Object, objType As Type, name As String, ParamArray values As Object()) As Object
        Try
            Return objType.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.InvokeMethod, Nothing, obj, values)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance function for """ & name & """", ex)
        End Try
    End Function

    Public Shared Function InvokeFunction(Of T)(obj As Object, objType As Type, name As String, ParamArray values As Object()) As T
        Return CType(InvokeFunction(obj, objType, name, values), T)
    End Function

    Public Shared Sub InvokeMethod(obj As Object, objType As Type, name As String, ParamArray values As Object())
        Try
            objType.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.InvokeMethod, Nothing, obj, values)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance method """ & name & """", ex)
        End Try
    End Sub

    Public Shared Function InvokePropertyGet(obj As Object, objtype As Type, name As String) As Object
        Try
            Return objtype.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.GetProperty, Nothing, obj, Nothing)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance property get """ & name & """", ex)
        End Try
    End Function

    Public Shared Function InvokePropertyGet(Of T)(obj As Object, objtype As Type, name As String) As T
        Try
            Return CType(objtype.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.GetProperty, Nothing, obj, Nothing), T)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance property get """ & name & """", ex)
        End Try
    End Function

    Public Shared Function InvokePropertyGet(obj As Object, objtype As Type, name As String, propertyArrayItem As Object) As Object
        Try
            Return objtype.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.GetProperty, Nothing, obj, New Object() {propertyArrayItem})
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance property get """ & name & """", ex)
        End Try
    End Function

    Public Shared Function InvokePropertyGet(Of T)(obj As Object, objtype As Type, name As String, propertyArrayItem As Object) As T
        Try
            Return CType(objtype.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.GetProperty, Nothing, obj, New Object() {propertyArrayItem}), T)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance property get """ & name & """", ex)
        End Try
    End Function

    Public Shared Sub InvokePropertySet(Of T)(obj As Object, objType As Type, name As String, value As T)
        Try
            objType.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.SetProperty, Nothing, obj, ConvertArgumentsToObjectArray(value))
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance property set """ & name & """", ex)
        End Try
    End Sub

    Public Shared Sub InvokePropertySet(Of T)(obj As Object, objType As Type, name As String, values As T())
        Try
            objType.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.SetProperty, Nothing, obj, ConvertArgumentsToObjectArray(values))
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance property set """ & name & """", ex)
        End Try
    End Sub

    Public Shared Function InvokeFieldGet(obj As Object, objtype As Type, name As String) As Object
        Try
            Return objtype.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.GetField, Nothing, obj, Nothing)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance field get """ & name & """", ex)
        End Try
    End Function

    Public Shared Function InvokeFieldGet(Of T)(obj As Object, objtype As Type, name As String) As T
        Try
            Return CType(objtype.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.GetField, Nothing, obj, Nothing), T)
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance field get """ & name & """", ex)
        End Try
    End Function

    Public Shared Sub InvokeFieldSet(Of T)(obj As Object, objType As Type, name As String, value As T)
        Try
            objType.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.SetField, Nothing, obj, ConvertArgumentsToObjectArray(value))
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance field set """ & name & """", ex)
        End Try
    End Sub

    Public Shared Sub InvokeFieldSet(Of T)(obj As Object, objType As Type, name As String, values As T())
        Try
            objType.InvokeMember(name, System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.SetField, Nothing, obj, ConvertArgumentsToObjectArray(values))
        Catch ex As Exception
            Throw New Exception("InvokeMember failed for public instance field set """ & name & """", ex)
        End Try
    End Sub
    Private Shared Function ConvertArgumentsToObjectArray(Of T)(value As T) As Object()
        Dim Args As Object()
        If value IsNot Nothing Then
            Dim ArgsList As New List(Of Object)
            ArgsList.Add(value)
            Args = ArgsList.ToArray
        Else
            Args = Nothing
        End If
        Return Args
    End Function

    Private Shared Function ConvertArgumentsToObjectArray(Of T)(values As T()) As Object()
        Dim Args As Object()
        If values IsNot Nothing Then
            Dim ArgsList As New List(Of Object)
            For MyCounter As Integer = 0 To values.Count - 1
                ArgsList.Add(values(MyCounter))
            Next
            Args = ArgsList.ToArray
        Else
            Args = Nothing
        End If
        Return Args
    End Function
End Class