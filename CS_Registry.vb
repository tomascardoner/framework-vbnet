Module CSM_Registry
    Private mRegistryConfigurationSubkeyName As String = "Software\" & My.Application.Info.CompanyName & "\" & My.Application.Info.ProductName

    '/////////////////////////////////////////////////////////////
    'USER VALUES
    '/////////////////////////////////////////////////////////////
    Friend Function LoadUserValue(ByVal SubKey As String, ByVal Name As String, ByVal DefaultValue As Object) As Object
        Dim RegistryKey As Microsoft.Win32.RegistryKey = Nothing

        Try
            RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(mRegistryConfigurationSubkeyName & CStr(IIf(SubKey = "", "", "\" & SubKey)), False)
            If Not RegistryKey Is Nothing Then
                LoadUserValue = RegistryKey.GetValue(Name, DefaultValue)
                RegistryKey.Close()
            Else
                LoadUserValue = DefaultValue
            End If
        Catch ex As Exception
            CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al obtener el Valor del Registro." & ControlChars.Cr & ControlChars.Cr & "Name: " & Name)
            If Not RegistryKey Is Nothing Then
                RegistryKey.Close()
            End If
            Return Nothing
        End Try
    End Function

    Friend Function SaveUserValue(ByVal SubKey As String, ByVal Name As String, ByVal Value As Object) As Boolean
        Dim RegistryKey As Microsoft.Win32.RegistryKey = Nothing

        Try
            RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(mRegistryConfigurationSubkeyName & CStr(IIf(SubKey = "", "", "\" & SubKey)))
            RegistryKey.SetValue(Name, Value)
            RegistryKey.Close()
            Return True

        Catch ex As Exception
            CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al guardar el Parámetro." & ControlChars.Cr & ControlChars.Cr & "Parameto: " & Name)
            If Not RegistryKey Is Nothing Then
                RegistryKey.Close()
            End If
            Return False
        End Try
    End Function

    '/////////////////////////////////////////////////////////////
    'APPLICATION VALUES
    '/////////////////////////////////////////////////////////////
    Friend Function LoadApplicationValue(ByVal SubKey As String, ByVal Name As String, ByVal DefaultValue As Object) As Object
        Dim RegistryKey As Microsoft.Win32.RegistryKey = Nothing

        Try
            RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(mRegistryConfigurationSubkeyName & CStr(IIf(SubKey = "", "", "\" & SubKey)), False)
            If Not RegistryKey Is Nothing Then
                LoadApplicationValue = RegistryKey.GetValue(Name, DefaultValue)
                RegistryKey.Close()
            Else
                LoadApplicationValue = DefaultValue
            End If
        Catch ex As Exception
            CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al obtener el Valor del Registro." & ControlChars.Cr & ControlChars.Cr & "Name: " & Name)
            If Not RegistryKey Is Nothing Then
                RegistryKey.Close()
            End If
            Return Nothing
        End Try
    End Function

    Friend Function SaveApplicationValue(ByVal SubKey As String, ByVal Name As String, ByVal Value As Object) As Boolean
        Dim RegistryKey As Microsoft.Win32.RegistryKey = Nothing

        Try
            RegistryKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(mRegistryConfigurationSubkeyName & CStr(IIf(SubKey = "", "", "\" & SubKey)))
            RegistryKey.SetValue(Name, Value)
            RegistryKey.Close()
            Return True
        Catch ex As Exception
            CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al guardar el Parámetro." & ControlChars.Cr & ControlChars.Cr & "Parameto: " & Name)
            If Not RegistryKey Is Nothing Then
                RegistryKey.Close()
            End If
            Return False
        End Try
    End Function

    '/////////////////////////////////////////////////////////////
    'GENERAL VALUES
    '/////////////////////////////////////////////////////////////
    Friend Function LoadValue(ByVal KeyName As String, ByVal Name As String, ByVal DefaultValue As Object) As Object
        Dim RegistryKey As Microsoft.Win32.RegistryKey = Nothing

        Try
            'RegistryKey = Microsoft.Win32.Registry.(KeyName, Name, DefaultValue)
            'If Not RegistryKey Is Nothing Then
            LoadValue = Microsoft.Win32.Registry.GetValue(KeyName, Name, DefaultValue)
            'RegistryKey.Close()
            'Else
            'LoadValue = DefaultValue
            'End If
        Catch ex As Exception
            CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al obtener el Valor del Registro." & ControlChars.Cr & ControlChars.Cr & "Name: " & Name)
            If Not RegistryKey Is Nothing Then
                RegistryKey.Close()
            End If
            Return Nothing
        End Try
    End Function

    Friend Function SaveValue(ByVal SubKey As String, ByVal Name As String, ByVal Value As Object) As Boolean
        Dim RegistryKey As Microsoft.Win32.RegistryKey = Nothing

        Try
            RegistryKey = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(mRegistryConfigurationSubkeyName & CStr(IIf(SubKey = "", "", "\" & SubKey)))
            RegistryKey.SetValue(Name, Value)
            RegistryKey.Close()
            Return True
        Catch ex As Exception
            CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al guardar el Parámetro." & ControlChars.Cr & ControlChars.Cr & "Parameto: " & Name)
            If Not RegistryKey Is Nothing Then
                RegistryKey.Close()
            End If
            Return False
        End Try
    End Function
End Module