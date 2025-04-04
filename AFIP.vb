Imports System.Text.RegularExpressions

Namespace CardonerSistemas

    Module Afip
        Friend Const CuitPrefijoMasculino As String = "20"
        Friend Const CuitPrefijoMasculinoAlternativo As String = "23"
        Friend Const CuitPrefijoFemenino As String = "27"
        Friend Const CuitPrefijoFemeninoAlternativo As String = "24"
        Friend Const CuitPrefijoPersonaJuridica As String = "30"
        Friend Const CuitPrefijoPersonaJuridicaAlternativo1 As String = "33"
        Friend Const CuitPrefijoPersonaJuridicaAlternativo2 As String = "34"

        ' Códigos QR en comprobantes
        Friend Const ComprobantesCodigoQRDataField As String = "{DATA}"
        Friend Const ComprobantesCodigoQRDataFieldFecha As String = "{FECHA}"
        Friend Const ComprobantesCodigoQRDataFieldCuit As String = "{CUIT}"
        Friend Const ComprobantesCodigoQRDataFieldPuntoVenta As String = "{PUNTOVENTA}"
        Friend Const ComprobantesCodigoQRDataFieldTipoComprobante As String = "{TIPOCOMPROBANTE}"
        Friend Const ComprobantesCodigoQRDataFieldNumeroComprobante As String = "{NUMEROCOMPROBANTE}"
        Friend Const ComprobantesCodigoQRDataFieldImporte As String = "{IMPORTE}"
        Friend Const ComprobantesCodigoQRDataFieldMoneda As String = "{MONEDA}"
        Friend Const ComprobantesCodigoQRDataFieldCotizacion As String = "{COTIZACION}"
        Friend Const ComprobantesCodigoQRDataFieldTipoDocumento As String = "{TIPODOCUMENTO}"
        Friend Const ComprobantesCodigoQRDataFieldNumeroDocumento As String = "{NUMERODOCUMENTO}"
        Friend Const ComprobantesCodigoQRDataFieldTipoCodigoAutorizacion As String = "{TIPOCODIGOAUTORIZACION}"
        Friend Const ComprobantesCodigoQRDataFieldCodigoAutorizacion As String = "{CODIGOAUTORIZACION}"

        Friend Enum TipoPersonas As Byte
            Femenino
            Masculino
            PersonaJurdica
        End Enum

        Friend Function ObtenerDigitoVerificadorCuit(ByVal Cuit As String) As Byte?
            Dim CuitLimpio As String
            Dim Total As Integer
            Dim Digito As Byte

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            Cuit = Cuit.Trim
            CuitLimpio = Regex.Replace(Cuit, "[^\d]", "")

            ' Verifico que el número tenga el formato correcto de:
            ' 10 dígitos consecutivos o 12 caracteres con guiones según el formato (99-99999999-)
            Select Case Cuit.Length
                Case 10
                    If CuitLimpio.Length < 10 Then
                        Return Nothing
                    End If
                Case 12
                    If Cuit.ElementAt(2) = "-" AndAlso Cuit.ElementAt(9) = "-" Then
                        If CuitLimpio.Length < 10 Then
                            Return Nothing
                        End If
                    Else
                        Return Nothing
                    End If
                Case Else
                    Return Nothing
            End Select
            Cuit = CuitLimpio

            ' Individualiza y multiplica los dígitos
            Total = CInt(Cuit.ElementAt(0).ToString) * 5
            Total += CInt(Cuit.ElementAt(1).ToString) * 4
            Total += CInt(Cuit.ElementAt(2).ToString) * 3
            Total += CInt(Cuit.ElementAt(3).ToString) * 2
            Total += CInt(Cuit.ElementAt(4).ToString) * 7
            Total += CInt(Cuit.ElementAt(5).ToString) * 6
            Total += CInt(Cuit.ElementAt(6).ToString) * 5
            Total += CInt(Cuit.ElementAt(7).ToString) * 4
            Total += CInt(Cuit.ElementAt(8).ToString) * 3
            Total += CInt(Cuit.ElementAt(9).ToString) * 2

            ' Calcula el dígito verificador
            Digito = CByte((11 - (Total Mod 11)) Mod 11)
            If Digito < 10 Then
                Return Digito
            Else
                Return Nothing
            End If
        End Function

        Friend Function ObtenerCuit(ByVal tipoPersona As TipoPersonas, ByVal numeroDocumento As String) As String
            Dim prefijo As String
            Dim prefijoAlternativo1 As String
            Dim prefijoAlternativo2 As String = ""
            Dim digitoVerificador As Byte?

            ' Selecciono el prefijo estándar
            Select Case tipoPersona
                Case TipoPersonas.Femenino
                    prefijo = CuitPrefijoFemenino
                    prefijoAlternativo1 = CuitPrefijoFemeninoAlternativo
                Case TipoPersonas.Masculino
                    prefijo = CuitPrefijoMasculino
                    prefijoAlternativo1 = CuitPrefijoMasculinoAlternativo
                Case TipoPersonas.PersonaJurdica
                    prefijo = CuitPrefijoPersonaJuridica
                    prefijoAlternativo1 = CuitPrefijoPersonaJuridicaAlternativo1
                    prefijoAlternativo2 = CuitPrefijoPersonaJuridicaAlternativo2
                Case Else
                    Return String.Empty
            End Select

            ' Verifico el largo del número de documento
            Select Case numeroDocumento.Length
                Case Is < 8
                    numeroDocumento = numeroDocumento.PadLeft(8, "0"c)
                Case > 8
                    numeroDocumento = numeroDocumento.Substring(0, 8)
            End Select

            ' Obtengo el dígito verificador
            digitoVerificador = Afip.ObtenerDigitoVerificadorCuit(prefijo & numeroDocumento)
            If Not digitoVerificador.HasValue Then
                ' No se pudo obtener, intentar con el prefijo alternativo
                digitoVerificador = CardonerSistemas.Afip.ObtenerDigitoVerificadorCuit(prefijoAlternativo1 & numeroDocumento)
                If Not digitoVerificador.HasValue Then
                    ' Tampoco se pudo obtener con el prefijo alternativo. Si es empresa, pruebo con el 2º alternativo
                    If tipoPersona = TipoPersonas.PersonaJurdica Then
                        digitoVerificador = CardonerSistemas.Afip.ObtenerDigitoVerificadorCuit(prefijoAlternativo2 & numeroDocumento)
                        If Not digitoVerificador.HasValue Then
                            ' Tampoco con el alternativo 2
                            Return String.Empty
                        End If
                    Else
                        Return String.Empty
                    End If
                End If
            End If

            ' Confirmo el CUIT verificándolo
            If CardonerSistemas.Afip.VerificarCuit(prefijo & numeroDocumento & digitoVerificador.ToString()) Then
                Return prefijo & numeroDocumento & digitoVerificador.ToString()
            Else
                Return String.Empty
            End If
        End Function

        Friend Function VerificarCuit(ByVal Cuit As String) As Boolean
            Dim CuitLimpio As String
            Dim Prefijo As String
            Dim DigitoVerificador As Byte?

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            Cuit = Cuit.Trim
            CuitLimpio = Regex.Replace(Cuit, "[^\d]", "")

            ' Verifico que el número tenga el formato correcto de:
            ' 11 dígitos consecutivos o 13 caracteres con guiones según el formato (99-99999999-9)
            Select Case Cuit.Length
                Case Is < 11
                    Return False
                Case 11
                    If CuitLimpio.Length < 11 Then
                        Return False
                    End If
                Case 13
                    If Cuit.ElementAt(2) = "-" AndAlso Cuit.ElementAt(9) = "-" Then
                        If CuitLimpio.Length < 11 Then
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Case Is > 13
                    Return False
            End Select
            Cuit = CuitLimpio

            Prefijo = Cuit.Substring(0, 2)
            If Prefijo = CuitPrefijoFemenino _
                OrElse Prefijo = CuitPrefijoFemeninoAlternativo _
                OrElse Prefijo = CuitPrefijoMasculino _
                OrElse Prefijo = CuitPrefijoMasculinoAlternativo _
                OrElse Prefijo = CuitPrefijoPersonaJuridica _
                OrElse Prefijo = CuitPrefijoPersonaJuridicaAlternativo1 _
                OrElse Prefijo = CuitPrefijoPersonaJuridicaAlternativo2 Then

                DigitoVerificador = ObtenerDigitoVerificadorCUIT(Cuit.Substring(0, 10))
                If DigitoVerificador Is Nothing Then
                    Return False
                Else
                    Return CBool(CByte(Cuit.Chars(10).ToString) = DigitoVerificador)
                End If
            Else
                Return False
            End If
        End Function

    End Module

End Namespace