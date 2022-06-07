Imports System.Reflection

Namespace CardonerSistemas

    Module Assembly

        ' <summary>
        ' Call the asssembly dynamically and execute a method
        '
        ' </summary>
        ' <param name="AssemblyName">Name of the Assembly to be loaded</param>
        ' <param name="className">Name of the class to be intantiated </param>
        ' <param name="methodName">Name of the method to be called</param>
        ' <param name="parameterForTheMethod">Parameters should be passed as object array</param>
        ' <returns>Returns as Generic object..</returns>

        Public Function Process(AssemblyName As String, className As String, methodName As String, parameterForTheMethod As Object()) As Object
            Dim returnObject As Object = Nothing
            Dim mi As MethodInfo = Nothing
            Dim ci As ConstructorInfo = Nothing
            Dim responder As Object = Nothing
            Dim Type As Type = Nothing
            Dim objectTypes As System.Type()
            Dim count As Integer = 0

            Try
                'Load the assembly and get it's information
                Type = System.Reflection.Assembly.LoadFrom(AssemblyName + ".dll").GetType(AssemblyName + "." + className)
                'Get the Passed parameter types to find the method type
                objectTypes = New System.Type(parameterForTheMethod.GetUpperBound(0) + 1)
                For Each objectParameter As Object In parameterForTheMethod
                    If objectParameter IsNot Nothing Then
                        objectTypes(count) = objectParameter.GetType()
                    End If
                    count += 1
                Next

                'Get the reference of the method
                mi = Type.GetMethod(methodName, objectTypes)
                ci = Type.GetConstructor(Type.EmptyTypes)
                responder = ci.Invoke(Nothing)
                'Invoke the method
                returnObject = mi.Invoke(responder, parameterForTheMethod)

            Catch ex As Exception
                Throw

            Finally
                mi = Nothing
                ci = Nothing
                responder = Nothing
                Type = Nothing
                objectTypes = Nothing
            End Try

            'Return the value as a generic object
            Return returnObject
        End Function

    End Module

End Namespace