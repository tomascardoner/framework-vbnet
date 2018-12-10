Imports System.Text.RegularExpressions

Module CS_Bank
    ' Formato de CBU
    'La CBU debe ser ingresada en 2 bloques:

    '      El 1º bloque contiene:
    '                • Banco (3 dígitos)
    '                • Dígito Verificador 1 (1 dígito)
    '                • Sucursal (3 dígitos)
    '                • Dígito Verificador 2 (1 digito)

    '      El 2º bloque contiene:
    '                • Cuenta (13 dígitos)
    '                • Dígito Verificador (1 dígito)

    Friend Function ObtenerDigitoVerificadorCBU_Bloque1(ByVal CBU_Bloque1 As String) As Byte?
        Dim CBU_Bloque1_Limpio As String
        Dim Total As Integer

        ' Limpio los espacios anterior y posterior que pudiera tener el string
        CBU_Bloque1 = CBU_Bloque1.Trim
        CBU_Bloque1_Limpio = Regex.Replace(CBU_Bloque1, "[^\d]", "")

        ' Verifico que el número tenga la longitud correcta
        If CBU_Bloque1.Length = 7 Then
            If CBU_Bloque1_Limpio.Length < 7 Then
                Return Nothing
            End If
        Else
            Return Nothing
        End If
        CBU_Bloque1 = CBU_Bloque1_Limpio

        ' Individualiza y multiplica los dígitos
        Total = CInt(CBU_Bloque1.ElementAt(0).ToString) * 7
        Total += CInt(CBU_Bloque1.ElementAt(1).ToString) * 1
        Total += CInt(CBU_Bloque1.ElementAt(2).ToString) * 3
        Total += CInt(CBU_Bloque1.ElementAt(3).ToString) * 9
        Total += CInt(CBU_Bloque1.ElementAt(4).ToString) * 7
        Total += CInt(CBU_Bloque1.ElementAt(5).ToString) * 1
        Total += CInt(CBU_Bloque1.ElementAt(6).ToString) * 3

        ' Calcula el dígito verificador
        Return CByte((10 - CByte(Total.ToString.Last.ToString)).ToString.Last.ToString)
    End Function

    Friend Function ObtenerDigitoVerificadorCBU_Bloque2(ByVal CBU_Bloque2 As String) As Byte?
        Dim CBU_Bloque2_Limpio As String
        Dim Total As Integer

        ' Limpio los espacios anterior y posterior que pudiera tener el string
        CBU_Bloque2 = CBU_Bloque2.Trim
        CBU_Bloque2_Limpio = Regex.Replace(CBU_Bloque2, "[^\d]", "")

        ' Verifico que el número tenga la longitud correcta
        If CBU_Bloque2.Length = 13 Then
            If CBU_Bloque2_Limpio.Length < 13 Then
                Return Nothing
            End If
        Else
            Return Nothing
        End If
        CBU_Bloque2 = CBU_Bloque2_Limpio

        ' Individualiza y multiplica los dígitos
        Total = CInt(CBU_Bloque2.ElementAt(0).ToString) * 3
        Total += CInt(CBU_Bloque2.ElementAt(1).ToString) * 9
        Total += CInt(CBU_Bloque2.ElementAt(2).ToString) * 7
        Total += CInt(CBU_Bloque2.ElementAt(3).ToString) * 1
        Total += CInt(CBU_Bloque2.ElementAt(4).ToString) * 3
        Total += CInt(CBU_Bloque2.ElementAt(5).ToString) * 9
        Total += CInt(CBU_Bloque2.ElementAt(6).ToString) * 7
        Total += CInt(CBU_Bloque2.ElementAt(7).ToString) * 1
        Total += CInt(CBU_Bloque2.ElementAt(8).ToString) * 3
        Total += CInt(CBU_Bloque2.ElementAt(9).ToString) * 9
        Total += CInt(CBU_Bloque2.ElementAt(10).ToString) * 7
        Total += CInt(CBU_Bloque2.ElementAt(11).ToString) * 1
        Total += CInt(CBU_Bloque2.ElementAt(12).ToString) * 3

        ' Calcula el dígito verificador
        Return CByte((10 - CByte(Total.ToString.Last.ToString)).ToString.Last.ToString)
    End Function

    ' Esta función verifica una CBU y devuelve:
    '   -1 si es correcta
    '    0 si es incorrecta por el formato
    '    1 si tiene un error en el primer bloque
    '    2 si tiene un error en el segundo bloque
    Friend Function VerificarCBU(ByVal CBU As String) As Short
        Dim CBU_Limpio As String
        Dim DigitoVerificador As Byte?

        ' Limpio los espacios anterior y posterior que pudiera tener el string
        CBU = CBU.Trim
        CBU_Limpio = Regex.Replace(CBU, "[^\d]", "")

        ' Verifico que el número tenga el formato correcto de:
        ' 22 dígitos consecutivos o 25 caracteres con guiones según el formato (0000000-0 0000000000000-0)
        Select Case CBU.Length
            Case 22
                If CBU_Limpio.Length < 22 Then
                    Return 0
                End If
            Case 25
                If CBU.ElementAt(7) = "-" And CBU.ElementAt(9) = " " And CBU.ElementAt(23) = "-" Then
                    If CBU_Limpio.Length < 22 Then
                        Return 0
                    End If
                Else
                    Return 0
                End If
            Case Else
                Return 0
        End Select
        CBU = CBU_Limpio

        ' Verifico el bloque 1
        DigitoVerificador = ObtenerDigitoVerificadorCBU_Bloque1(CBU.Substring(0, 7))
        If DigitoVerificador Is Nothing Then
            Return 0
        Else
            If CByte(CBU.Chars(7).ToString) <> DigitoVerificador Then
                Return 1
            End If
        End If

        ' Verifico el bloque 2
        DigitoVerificador = ObtenerDigitoVerificadorCBU_Bloque2(CBU.Substring(8, 13))
        If DigitoVerificador Is Nothing Then
            Return 0
        Else
            If CByte(CBU.Chars(21).ToString) <> DigitoVerificador Then
                Return 2
            End If
        End If

        ' CBU correcto
        Return -1
    End Function

    Friend Function CBUCorrecta(ByVal CBU As String) As Boolean
        Return CBool(VerificarCBU(CBU) = -1)
    End Function
End Module