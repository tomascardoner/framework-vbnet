Module CS_Parameter_System

    Friend Function GetString(idParametro As String, Optional defaultValue As String = Nothing) As String
        Dim parametroActual As Parametro

        parametroActual = pParametros.Find(Function(p) p.IDParametro.TrimEnd = idParametro)
        If parametroActual Is Nothing OrElse parametroActual.Texto Is Nothing Then
            Return defaultValue
        Else
            Return parametroActual.Texto.Trim
        End If
    End Function

    Friend Sub SetString(idParametro As String, value As String)
        Dim parametroNuevo As New Parametro With {.IDParametro = idParametro, .Texto = value}

        ' Lo guardo en la base de datos
        Parametros.SaveParameter(parametroNuevo)

        ' Actualizo la colección en memoria
        Dim parametroActual As Parametro = pParametros.Find(Function(p) p.IDParametro.TrimEnd = idParametro)
        If parametroActual Is Nothing Then
            pParametros.Add(parametroNuevo)
        Else
            parametroActual.Texto = value
        End If
        parametroNuevo = Nothing
    End Sub

    Friend Function GetIntegerAsByte(idParametro As String, Optional defaultValue As Byte = Nothing) As Byte
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = idParametro)
        If ParametroCurrent Is Nothing OrElse Not ParametroCurrent.NumeroEntero.HasValue Then
            ParametroCurrent = Nothing
            Return defaultValue
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

    Friend Function SetIntegerAsInteger(ByVal IDParametro As String, ByVal Value As Integer) As Boolean
        Dim ParametroCurrent As Parametro

        ParametroCurrent = pParametros.Find(Function(param) param.IDParametro.TrimEnd = IDParametro)
        If ParametroCurrent Is Nothing Then
            ParametroCurrent = New Parametro With {
                .IDParametro = IDParametro
            }
        End If
        ParametroCurrent.NumeroEntero = Value
        Return Parametros.SaveParameter(ParametroCurrent)
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