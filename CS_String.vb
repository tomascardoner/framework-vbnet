Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Globalization
Imports System.Runtime.CompilerServices


Friend Module CS_String

    ''' <summary>
    ''' Gets a sub-string given it's zero-based order position and separator
    ''' </summary>
    ''' <param name="value"></param>
    ''' <param name="orderPosition"></param>
    ''' <param name="separator"></param>
    ''' <returns></returns>
    <Extension()>
    Friend Function GetSubString(value As String, orderPosition As Integer, ByVal separator As Char) As String
        Dim aArray As String()

        aArray = value.Split(separator)
        If orderPosition <= aArray.Length Then
            Return aArray(orderPosition - 1)
        Else
            Return String.Empty
        End If
    End Function

    Friend Function GetParameterValueFromString(ByVal FullString As String, ByVal ParameterName As String, Optional ByVal DefaultValue As String = "", Optional ByVal ParametersSeparator As Char = ";"c) As String
        Dim CParameters As String()
        Dim ParameterFull As String
        Dim ParameterValue As String = Nothing

        CParameters = FullString.Split(ParametersSeparator)
        For Each ParameterFull In CParameters
            If ParameterFull.Substring(0, ParameterName.Length) = ParameterName Then
                ParameterValue = ParameterFull.Substring(ParameterName.Length + 1)
                Exit For
            End If
        Next
        If ParameterValue Is Nothing Then
            Return DefaultValue
        Else
            Return ParameterValue
        End If
    End Function

    Friend Function RemoveFirstSubString(ByVal MainString As String, ByVal SubStringSeparator As String) As String
        Return MainString.Remove(0, MainString.IndexOf(SubStringSeparator) + 1)
    End Function

    Friend Function RemoveLastSubString(ByVal MainString As String, ByVal SubStringSeparator As String) As String
        Dim SubStringIndex As Integer

        SubStringIndex = MainString.LastIndexOf(SubStringSeparator)
        If SubStringIndex = -1 Then
            Return String.Empty
        Else
            Return MainString.Remove(SubStringIndex)
        End If
    End Function

    Friend Function CountSubString(ByVal MainString As String, ByVal SubStringSeparator As String) As Integer
        Dim aArray As String()

        aArray = Split(MainString, SubStringSeparator)
        Return aArray.Length
    End Function

    Friend Function GetBooleanString(ByVal Value As Boolean) As String
        Return IIf(Value, My.Resources.STRING_YES, My.Resources.STRING_NO).ToString
    End Function

    Friend Function GetBooleanValueFromString(ByVal Value As String) As Boolean
        Dim testAgainst As String() = {"yes", "true", "si", "sí", "verdadero"}

        Return testAgainst.Contains(Value.ToLowerInvariant)
    End Function

    Friend Function RemoveInvalidCharsByAllowed(ByVal value As String, ByVal allowedChars As String) As String
        Dim charIndex As Integer
        Dim charBeingAnalyzed As Char
        Dim sb As New StringBuilder()

        For charIndex = 0 To value.Length - 1
            charBeingAnalyzed = value.Chars(charIndex)
            If allowedChars.Contains(charBeingAnalyzed) Then
                sb.Append(charBeingAnalyzed)
            End If
        Next charIndex
        Return sb.ToString()
    End Function

    Friend Function RemoveInvalidCharsByNotAllowed(ByVal value As String, ByVal notAllowedChars As String) As String
        Dim charIndex As Integer
        Dim charBeingAnalyzed As Char
        Dim sb As New StringBuilder()

        For charIndex = 0 To value.Length - 1
            charBeingAnalyzed = value.Chars(charIndex)
            If Not notAllowedChars.Contains(charBeingAnalyzed) Then
                sb.Append(charBeingAnalyzed)
            End If
        Next charIndex
        Return sb.ToString()
    End Function

    Friend Function RemoveInvalidSpaces(ByVal Value As String, Optional TrimValue As Boolean = True) As String
        Dim CleanedString As String

        If TrimValue Then
            CleanedString = Value.Trim
        Else
            CleanedString = Value
        End If
        Do While Value.Contains("  ")
            CleanedString = CleanedString.Replace("  ", " ")
        Loop
        Return CleanedString
    End Function

    Friend Function RemoveNullChars(ByVal Value As String) As String
        Return Value.Replace(vbNullChar, String.Empty)
    End Function

    <Extension()>
    Friend Function RemoveNotNumbers(ByVal value As String) As String
        If value IsNot Nothing Then
            Return Regex.Replace(value, "[^0-9]", String.Empty)
        Else
            Return value
        End If
    End Function

    <Extension()>
    Friend Function IsAllDigits(ByVal value As String) As Boolean
        If value IsNot Nothing Then
            Return (value.Length = RemoveNotNumbers(value).Length)
        Else
            Return False
        End If
    End Function

    <Extension()>
    Friend Function RemoveDiacritics(ByVal value As String) As String
        Dim NormalizedString As String = value.Normalize(NormalizationForm.FormD)
        Dim StringBuilder As New StringBuilder

        For Each character As Char In NormalizedString
            Dim UnicodeCategoryCurrent As UnicodeCategory = CharUnicodeInfo.GetUnicodeCategory(character)
            If UnicodeCategoryCurrent <> UnicodeCategory.NonSpacingMark Then
                StringBuilder.Append(character)
            End If
        Next

        Return StringBuilder.ToString().Normalize(NormalizationForm.FormC)
    End Function

    <Extension()>
    Friend Function TrimAndReduce(ByVal value As String) As String
        Return ConvertWhitespacesToSingleSpaces(value).Trim()
    End Function

    <Extension()>
    Friend Function ConvertWhitespacesToSingleSpaces(ByVal value As String) As String
        Return Regex.Replace(value, "\s+", " ")
    End Function

    <Extension()>
    Friend Function Truncate(value As String, maxLength As Integer) As String
        If String.IsNullOrEmpty(value) OrElse value.Length <= maxLength Then
            Return value
        Else
            Return value.Substring(0, maxLength)
        End If
    End Function

    ''' <summary>
    ''' Converts the specified string to title case (except for words that are entirely in uppercase, which are considered to be acronyms).
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    <Extension()>
    Friend Function ToTitleCase(value As String) As String
        Return Application.CurrentCulture.TextInfo.ToTitleCase(value)
    End Function

    ''' <summary>
    ''' Converts the specified string to title case (also the words that are entirely in uppercase).
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    <Extension()>
    Friend Function ToTitleCaseAll(value As String) As String
        Return Application.CurrentCulture.TextInfo.ToTitleCase(value.ToLower())
    End Function

    Friend Function StripHtml(ByVal value As String) As String
        Return Regex.Replace(value, "<(.|\n)*?>", String.Empty)
    End Function

    <Extension>
    Friend Function GetEnclosedSubString(value As String, openChar As Char, closeChar As Char) As String
        Dim startIndex As Integer
        Dim endIndex As Integer

        startIndex = value.IndexOf(openChar)
        If startIndex > -1 Then
            endIndex = value.IndexOf(closeChar, startIndex + 1)
            If endIndex > -1 AndAlso (endIndex - startIndex - 1) > 0 Then
                Return value.Substring(startIndex + 1, endIndex - startIndex - 1)
            End If
        End If

        Return String.Empty
    End Function


    Friend Function GetExtends(ByRef graphicObject As System.Drawing.Graphics, text As String, font As System.Drawing.Font) As Integer
        Dim size As SizeF = graphicObject.MeasureString(text, font)

        Return CInt(Math.Ceiling(size.Width))
    End Function

End Module