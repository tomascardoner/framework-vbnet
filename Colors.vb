Imports System.Text.RegularExpressions

Namespace CardonerSistemas

    Module Colors

#Region "Hexadecimal validation"

        ' Hex RGB
        Friend Const RegExHexadecimalRgbSingleDigits As String = "^#(?:[0-9a-fA-F]{3})$"
        Friend Const RegExHexadecimalRgbDoubleDigits As String = "^#(?:[0-9a-fA-F]{6})$"
        Friend Const RegExHexadecimalRgbBothDigits As String = "^#(?:[0-9a-fA-F]{3,6})$"

        ' Hex ARGB
        Friend Const RegExHexadecimalArgbSingleDigits As String = "^#(?:[0-9a-fA-F]{4})$"
        Friend Const RegExHexadecimalArgbDoubleDigits As String = "^#(?:[0-9a-fA-F]{8})$"
        Friend Const RegExHexadecimalArgbBothDigits As String = "^#(?:[0-9a-fA-F]{4,8})$"

        ' Hex RGB or ARGB
        Friend Const RegExHexadecimalRgbOrArgbSingleDigits As String = "^#(?:[0-9a-fA-F]{3,4})$"
        Friend Const RegExHexadecimalRgbOrArgbDoubleDigits As String = "^#(?:[0-9a-fA-F]{6,8})$"
        Friend Const RegExHexadecimalRgbOrArgbBothDigits As String = "^#(?:[0-9a-fA-F]{3,4,6,8})$"

        Private Function IsValidHexColor(value As String, evaluateExpression As String) As Boolean
            If String.IsNullOrEmpty(value) Then
                Return False
            End If

            Try
                Return Regex.IsMatch(value, evaluateExpression, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250))
            Catch ex As Exception
                Return False
            End Try
        End Function

#End Region

#Region "Serialization Get"

        Friend Function GetFromHexString(value As String, ByRef color As Color, Optional evaluateExpression As String = RegExHexadecimalRgbOrArgbBothDigits, Optional valueDefault As String = Nothing) As Boolean
            If String.IsNullOrWhiteSpace(value) AndAlso Not String.IsNullOrWhiteSpace(valueDefault) Then
                value = valueDefault
            End If

            If String.IsNullOrWhiteSpace(value) Then
                value = value.Trim()
                If value.StartsWith("#") Then
                    If IsValidHexColor(value, evaluateExpression) Then
                        Try
                            color = ColorTranslator.FromHtml(value)
                            Return True
                        Catch ex As Exception
                            ErrorHandler.ProcessError(ex, "Error al convertir el valor en un color.")
                        End Try
                    End If
                End If
            End If
            Return False
        End Function

        Friend Function GetFromNameString(value As String, ByRef color As Color, Optional valueDefault As String = Nothing) As Boolean
            If String.IsNullOrWhiteSpace(value) AndAlso Not String.IsNullOrWhiteSpace(valueDefault) Then
                value = valueDefault
            End If

            If Not String.IsNullOrWhiteSpace(value) Then
                value = value.Trim()
                Dim namedColor As Color
                namedColor = Color.FromName(value)
                If (CShort(namedColor.A) + namedColor.R + namedColor.G + namedColor.B > 0) Then
                    color = namedColor
                    Return True
                End If
            End If
            Return False
        End Function

        Friend Function GetFromHexOrNameString(value As String, ByRef color As Color, Optional evaluateExpression As String = RegExHexadecimalRgbOrArgbBothDigits, Optional valueDefault As String = Nothing) As Boolean
            If String.IsNullOrWhiteSpace(value) AndAlso Not String.IsNullOrWhiteSpace(valueDefault) Then
                value = valueDefault
            End If

            If Not String.IsNullOrWhiteSpace(value) Then
                value = value.Trim()
                If value.StartsWith("#") Then
                    Return GetFromHexString(value, color, evaluateExpression, valueDefault)
                Else
                    Return GetFromNameString(value, color, valueDefault)
                End If
            End If
            Return False
        End Function

#End Region

#Region "Serialization Set"

        Friend Function SetToHexRgbString(color As Color) As String
            If color.IsEmpty Then
                Return String.Empty
            Else
                Return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2")
            End If
        End Function

        Friend Function SetToHexArgbString(color As Color) As String
            If color.IsEmpty Then
                Return String.Empty
            Else
                Return "#" + color.A.ToString("X2") + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2")
            End If
        End Function

        Friend Function SetToNamedOrHexRgbString(color As Color) As String
            If color.IsEmpty Then
                Return String.Empty
            Else
                If color.IsKnownColor Then
                    Return color.Name
                Else
                    Return SetToHexRgbString(color)
                End If
            End If
        End Function

        Friend Function SetToNamedOrHexArgbString(color As Color) As String
            If color.IsEmpty Then
                Return String.Empty
            Else
                If color.IsKnownColor Then
                    Return color.Name
                Else
                    Return SetToHexArgbString(color)
                End If
            End If
        End Function

#End Region

#Region "Assignation"

        Private Function SetColor(newColor As Color?, defaultColor As Color) As Color
            If newColor.HasValue Then
                Return newColor.Value
            Else
                Return defaultColor
            End If
        End Function

#End Region

    End Module

End Namespace