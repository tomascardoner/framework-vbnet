Imports System.Text.RegularExpressions

Namespace CardonerSistemas

    Module AFIP
        Friend Const CUIT_PEFIJO_MASCULINO As String = "20"
        Friend Const CUIT_PEFIJO_MASCULINO_ALTERNATIVO As String = "23"
        Friend Const CUIT_PEFIJO_FEMENINO As String = "27"
        Friend Const CUIT_PEFIJO_FEMENINO_ALTERNATIVO As String = "24"
        Friend Const CUIT_PEFIJO_PERSONAJURIDICA As String = "30"
        Friend Const CUIT_PEFIJO_PERSONAJURIDICA_ALTERNATIVO1 As String = "33"
        Friend Const CUIT_PEFIJO_PERSONAJURIDICA_ALTERNATIVO2 As String = "34"

        Friend Enum TipoPersonas As Byte
            Femenino
            Masculino
            PersonaJurdica
        End Enum

        Friend Function ObtenerDigitoVerificadorCUIT(ByVal CUIT As String) As Byte?
            Dim CUITLimpio As String
            Dim Total As Integer
            Dim Digito As Byte

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            CUIT = CUIT.Trim
            CUITLimpio = Regex.Replace(CUIT, "[^\d]", "")

            ' Verifico que el número tenga el formato correcto de:
            ' 10 dígitos consecutivos o 12 caracteres con guiones según el formato (99-99999999-)
            Select Case CUIT.Length
                Case 10
                    If CUITLimpio.Length < 10 Then
                        Return Nothing
                    End If
                Case 12
                    If CUIT.ElementAt(2) = "-" And CUIT.ElementAt(9) = "-" Then
                        If CUITLimpio.Length < 10 Then
                            Return Nothing
                        End If
                    Else
                        Return Nothing
                    End If
                Case Else
                    Return Nothing
            End Select
            CUIT = CUITLimpio

            ' Individualiza y multiplica los dígitos
            Total = CInt(CUIT.ElementAt(0).ToString) * 5
            Total += CInt(CUIT.ElementAt(1).ToString) * 4
            Total += CInt(CUIT.ElementAt(2).ToString) * 3
            Total += CInt(CUIT.ElementAt(3).ToString) * 2
            Total += CInt(CUIT.ElementAt(4).ToString) * 7
            Total += CInt(CUIT.ElementAt(5).ToString) * 6
            Total += CInt(CUIT.ElementAt(6).ToString) * 5
            Total += CInt(CUIT.ElementAt(7).ToString) * 4
            Total += CInt(CUIT.ElementAt(8).ToString) * 3
            Total += CInt(CUIT.ElementAt(9).ToString) * 2

            ' Calcula el dígito verificador
            Digito = CByte((11 - (Total Mod 11)) Mod 11)
            If Digito < 10 Then
                Return Digito
            Else
                Return Nothing
            End If
        End Function

        Friend Function ObtenerCUIT(ByVal tipoPersona As TipoPersonas, ByVal numeroDocumento As String) As String
            Dim prefijo As String
            Dim prefijoAlternativo1 As String
            Dim prefijoAlternativo2 As String
            Dim digitoVerificador As Byte?

            ' Selecciono el prefijo estándar
            Select Case tipoPersona
                Case TipoPersonas.Femenino
                    prefijo = CUIT_PEFIJO_FEMENINO
                    prefijoAlternativo1 = CUIT_PEFIJO_FEMENINO_ALTERNATIVO
                Case TipoPersonas.Masculino
                    prefijo = CUIT_PEFIJO_MASCULINO
                    prefijoAlternativo1 = CUIT_PEFIJO_MASCULINO_ALTERNATIVO
                Case TipoPersonas.PersonaJurdica
                    prefijo = CUIT_PEFIJO_PERSONAJURIDICA
                    prefijoAlternativo1 = CUIT_PEFIJO_PERSONAJURIDICA_ALTERNATIVO1
                    prefijoAlternativo2 = CUIT_PEFIJO_PERSONAJURIDICA_ALTERNATIVO2
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
            digitoVerificador = CardonerSistemas.AFIP.ObtenerDigitoVerificadorCUIT(prefijo + numeroDocumento)
            If Not digitoVerificador.HasValue Then
                ' No se pudo obtener, intentar con el prefijo alternativo
                digitoVerificador = CardonerSistemas.AFIP.ObtenerDigitoVerificadorCUIT(prefijoAlternativo1 + numeroDocumento)
                If Not digitoVerificador.HasValue Then
                    ' Tampoco se pudo obtener con el prefijo alternativo. Si es empresa, pruebo con el 2º alternativo
                    If tipoPersona = TipoPersonas.PersonaJurdica Then
                        digitoVerificador = CardonerSistemas.AFIP.ObtenerDigitoVerificadorCUIT(prefijoAlternativo2 + numeroDocumento)
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
            If CardonerSistemas.AFIP.VerificarCUIT(prefijo + numeroDocumento + digitoVerificador.ToString()) Then
                Return prefijo & numeroDocumento & digitoVerificador.ToString()
            Else
                Return String.Empty
            End If
        End Function


        Friend Function VerificarCUIT(ByVal CUIT As String) As Boolean
            Dim CUITLimpio As String
            Dim Prefijo As String
            Dim DigitoVerificador As Byte?

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            CUIT = CUIT.Trim
            CUITLimpio = Regex.Replace(CUIT, "[^\d]", "")

            ' Verifico que el número tenga el formato correcto de:
            ' 11 dígitos consecutivos o 13 caracteres con guiones según el formato (99-99999999-9)
            Select Case CUIT.Length
                Case Is < 11
                    Return False
                Case 11
                    If CUITLimpio.Length < 11 Then
                        Return False
                    End If
                Case 13
                    If CUIT.ElementAt(2) = "-" And CUIT.ElementAt(9) = "-" Then
                        If CUITLimpio.Length < 11 Then
                            Return False
                        End If
                    Else
                        Return False
                    End If
                Case Is > 13
                    Return False
            End Select
            CUIT = CUITLimpio

            Prefijo = CUIT.Substring(0, 2)
            If Prefijo = CUIT_PEFIJO_FEMENINO _
                Or Prefijo = CUIT_PEFIJO_FEMENINO_ALTERNATIVO _
                Or Prefijo = CUIT_PEFIJO_MASCULINO _
                Or Prefijo = CUIT_PEFIJO_MASCULINO_ALTERNATIVO _
                Or Prefijo = CUIT_PEFIJO_PERSONAJURIDICA _
                Or Prefijo = CUIT_PEFIJO_PERSONAJURIDICA_ALTERNATIVO1 _
                Or Prefijo = CUIT_PEFIJO_PERSONAJURIDICA_ALTERNATIVO2 Then

                DigitoVerificador = ObtenerDigitoVerificadorCUIT(CUIT.Substring(0, 10))
                If DigitoVerificador Is Nothing Then
                    Return False
                Else
                    Return CBool(CByte(CUIT.Chars(10).ToString) = DigitoVerificador)
                End If
            Else
                Return False
            End If
        End Function

    End Module

End Namespace