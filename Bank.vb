Imports System.Text.RegularExpressions

Namespace CardonerSistemas

    Friend Module Bank
        ' Formato de CBU
        ' La CBU debe ser ingresada en 2 bloques:

        ' El 1º bloque contiene:
        '      • Banco (3 dígitos)
        '      • Dígito Verificador 1 (1 dígito)
        '      • Sucursal (3 dígitos)
        '      • Dígito Verificador 2 (1 digito)

        ' El 2º bloque contiene:
        '      • Cuenta (13 dígitos)
        '      • Dígito Verificador (1 dígito)

        Friend Function ObtenerDigitoVerificadorCbuBloque1(cbuBloque1 As String) As Byte?
            Dim cbuBloque1Limpio As String
            Dim total As Integer

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            cbuBloque1 = cbuBloque1.Trim()
            cbuBloque1Limpio = Regex.Replace(cbuBloque1, "[^\d]", "")

            ' Verifico que el número tenga la longitud correcta
            If cbuBloque1.Length = 7 Then
                If cbuBloque1Limpio.Length < 7 Then
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
            cbuBloque1 = cbuBloque1Limpio

            ' Individualiza y multiplica los dígitos
            total = Convert.ToInt16(cbuBloque1(0).ToString()) * 7
            total += Convert.ToInt16(cbuBloque1(1).ToString()) * 1
            total += Convert.ToInt16(cbuBloque1(2).ToString()) * 3
            total += Convert.ToInt16(cbuBloque1(3).ToString()) * 9
            total += Convert.ToInt16(cbuBloque1(4).ToString()) * 7
            total += Convert.ToInt16(cbuBloque1(5).ToString()) * 1
            total += Convert.ToInt16(cbuBloque1(6).ToString()) * 3

            ' Calcula el dígito verificador
            Dim digitoVerificador As Byte
            digitoVerificador = CByte(10 - Convert.ToByte(total.ToString().Last().ToString()))

            If digitoVerificador = 10 Then
                Return 0
            Else
                Return digitoVerificador
            End If
        End Function

        Friend Function ObtenerDigitoVerificadorCbuBloque2(cbuBloque2 As String) As Byte?
            Dim cbuBloque2Limpio As String
            Dim total As Integer

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            cbuBloque2 = cbuBloque2.Trim()
            cbuBloque2Limpio = Regex.Replace(cbuBloque2, "[^\d]", "")

            ' Verifico que el número tenga la longitud correcta
            If cbuBloque2.Length = 13 Then
                If cbuBloque2Limpio.Length < 13 Then
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
            cbuBloque2 = cbuBloque2Limpio

            ' Individualiza y multiplica los dígitos
            total = Convert.ToInt16(cbuBloque2(0).ToString()) * 3
            total += Convert.ToInt16(cbuBloque2(1).ToString()) * 9
            total += Convert.ToInt16(cbuBloque2(2).ToString()) * 7
            total += Convert.ToInt16(cbuBloque2(3).ToString()) * 1
            total += Convert.ToInt16(cbuBloque2(4).ToString()) * 3
            total += Convert.ToInt16(cbuBloque2(5).ToString()) * 9
            total += Convert.ToInt16(cbuBloque2(6).ToString()) * 7
            total += Convert.ToInt16(cbuBloque2(7).ToString()) * 1
            total += Convert.ToInt16(cbuBloque2(8).ToString()) * 3
            total += Convert.ToInt16(cbuBloque2(9).ToString()) * 9
            total += Convert.ToInt16(cbuBloque2(10).ToString()) * 7
            total += Convert.ToInt16(cbuBloque2(11).ToString()) * 1
            total += Convert.ToInt16(cbuBloque2(12).ToString()) * 3

            ' Calcula el dígito verificador
            Dim digitoVerificador As Byte
            digitoVerificador = CByte(10 - Convert.ToByte(total.ToString().Last().ToString()))

            If digitoVerificador = 10 Then
                Return 0
            Else
                Return digitoVerificador
            End If
        End Function

        ' Esta función verifica una CBU y devuelve:
        '   -1 si es correcta
        '    0 si es incorrecta por el formato
        '    1 si tiene un error en el primer bloque
        '    2 si tiene un error en el segundo bloque
        Friend Function VerificarCBU(cbu As String) As Short
            Dim cbuLimpio As String
            Dim digitoVerificador As Byte?

            ' Limpio los espacios anterior y posterior que pudiera tener el string
            cbu = cbu.Trim()
            cbuLimpio = Regex.Replace(cbu, "[^\d]", "")

            ' Verifico que el número tenga el formato correcto de:
            ' 22 dígitos consecutivos o 25 caracteres con guiones según el formato (0000000-0 0000000000000-0)
            Select Case cbu.Length
                Case 22
                    If cbuLimpio.Length < 22 Then
                        Return 0
                    End If
                Case 25
                    If cbu(7) = "-" AndAlso cbu(9) = " " AndAlso cbu(23) = "-" Then
                        If cbuLimpio.Length < 22 Then
                            Return 0
                        End If
                    Else
                        Return 0
                    End If
                Case Else
                    Return 0
            End Select
            cbu = cbuLimpio

            ' Verifico el bloque 1
            digitoVerificador = ObtenerDigitoVerificadorCbuBloque1(cbu.Substring(0, 7))
            If digitoVerificador Is Nothing Then
                Return 0
            Else
                If Convert.ToByte(cbu(7).ToString()) <> digitoVerificador Then
                    Return 1
                End If
            End If

            ' Verifico el bloque 2
            digitoVerificador = ObtenerDigitoVerificadorCbuBloque2(cbu.Substring(8, 13))
            If digitoVerificador Is Nothing Then
                Return 0
            Else
                If Convert.ToByte(cbu(21).ToString()) <> digitoVerificador Then
                    Return 2
                End If
            End If

            ' CBU correcto
            Return -1
        End Function

        Friend Function CBUCorrecta(cbu As String) As Boolean
            Return (VerificarCBU(cbu) = -1)
        End Function

    End Module
    
End Namespace