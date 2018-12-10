Module CS_Parameter_User
    Friend Function UsuarioGetString(ByVal IDParametro As String, Optional ByVal DefaultValue As String = Nothing) As String
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse ParametroCurrent.Texto Is Nothing Then
            Return DefaultValue
        Else
            Return ParametroCurrent.Texto.Trim
        End If
    End Function

    Friend Function UsuarioGetIntegerAsByte(ByVal IDParametro As String, Optional ByVal DefaultValue As Byte = Nothing) As Byte
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            ParametroCurrent = Nothing
            Return DefaultValue
        Else
            Return CByte(ParametroCurrent.NumeroEntero.Value)
        End If
    End Function

    Friend Function UsuarioGetIntegerAsShort(ByVal IDParametro As String, Optional ByVal DefaultValue As Short = Nothing) As Short
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            ParametroCurrent = Nothing
            Return DefaultValue
        Else
            Return CShort(ParametroCurrent.NumeroEntero.Value)
        End If
    End Function

    Friend Function UsuarioGetIntegerAsInteger(ByVal IDParametro As String, Optional ByVal DefaultValue As Integer = Nothing) As Integer
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            Return DefaultValue
        Else
            Return ParametroCurrent.NumeroEntero.Value
        End If
    End Function

    Friend Function UsuarioGetDecimal(ByVal IDParametro As String, Optional ByVal DefaultValue As Decimal = Nothing) As Decimal
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroDecimal.HasValue Then
            Return DefaultValue
        Else
            Return ParametroCurrent.NumeroDecimal.Value
        End If
    End Function

    Friend Function UsuarioGetDate(ByVal IDParametro As String, Optional ByVal DefaultValue As Date = Nothing) As Date?
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse ParametroCurrent.FechaHora Is Nothing Then
            Return DefaultValue
        Else
            Return ParametroCurrent.FechaHora
        End If
    End Function

    Friend Function UsuarioGetBoolean(ByVal IDParametro As String, Optional ByVal DefaultValue As Boolean? = Nothing) As Boolean?
        Dim ParametroCurrent As UsuarioParametro

        ParametroCurrent = pUsuarioParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse ParametroCurrent.SiNo Is Nothing Then
            Return DefaultValue
        Else
            Return ParametroCurrent.SiNo
        End If
    End Function

    Friend Sub UsuarioSaveBoolean(ByVal IDParametro As String, ByVal Value As Boolean)

    End Sub

End Module
