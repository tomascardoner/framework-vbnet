Imports System.Text.RegularExpressions

Module CS_AFIP
    Friend Function ObtenerDigitoVerificadorCUIT(ByVal CUIT As String) As Byte?
        Dim CUITLimpio As String
        Dim Total As Integer

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
        Return CByte((11 - (Total Mod 11)) Mod 11)
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
        If Prefijo = "20" Or Prefijo = "23" Or Prefijo = "24" Or Prefijo = "27" Or Prefijo = "30" Or Prefijo = "33" Or Prefijo = "34" Then
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