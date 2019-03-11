Module CS_Parameter_System

    Friend Function GetString(ByVal IDParametro As String, Optional ByVal DefaultValue As String = Nothing) As String
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse ParametroCurrent.Texto Is Nothing Then
            Return DefaultValue
        Else
            Return ParametroCurrent.Texto.Trim
        End If
    End Function

    Friend Function GetIntegerAsByte(ByVal IDParametro As String, Optional ByVal DefaultValue As Byte = Nothing) As Byte
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            ParametroCurrent = Nothing
            Return DefaultValue
        Else
            Return CByte(ParametroCurrent.NumeroEntero.Value)
        End If
    End Function

    Friend Function GetIntegerAsShort(ByVal IDParametro As String, Optional ByVal DefaultValue As Short = Nothing) As Short
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            ParametroCurrent = Nothing
            Return DefaultValue
        Else
            Return CShort(ParametroCurrent.NumeroEntero.Value)
        End If
    End Function

    Friend Function GetIntegerAsInteger(ByVal IDParametro As String, Optional ByVal DefaultValue As Integer = Nothing) As Integer
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            Return DefaultValue
        Else
            Return ParametroCurrent.NumeroEntero.Value
        End If
    End Function

    Friend Function GetDecimal(ByVal IDParametro As String, Optional ByVal DefaultValue As Decimal = Nothing) As Decimal
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroDecimal.HasValue Then
            Return DefaultValue
        Else
            Return ParametroCurrent.NumeroDecimal.Value
        End If
    End Function

    Friend Function GetDate(ByVal IDParametro As String, Optional ByVal DefaultValue As Date = Nothing) As Date?
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse ParametroCurrent.FechaHora Is Nothing Then
            Return DefaultValue
        Else
            Return ParametroCurrent.FechaHora
        End If
    End Function

    Friend Function GetBoolean(ByVal IDParametro As String, Optional ByVal DefaultValue As Boolean? = Nothing) As Boolean?
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing OrElse ParametroCurrent.SiNo Is Nothing Then
            Return DefaultValue
        Else
            Return ParametroCurrent.SiNo
        End If
    End Function

End Module