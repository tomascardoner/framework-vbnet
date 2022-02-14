Imports System.Text
Imports System.Text.RegularExpressions

Module CS_ValueTranslation

#Region "De Objectos a Controles - TextBox (standard)"

    Friend Function FromObjectStringToControlTextBox(ByVal ObjectValue As String) As String
        If String.IsNullOrEmpty(ObjectValue) Then
            Return ""
        Else
            Return ObjectValue
        End If
    End Function

    Friend Function FromObjectMoneyToControlTextBox(ByVal ObjectValue As Decimal?) As String
        If ObjectValue.HasValue Then
            Return FormatCurrency(ObjectValue.Value)
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectIntegerToControlTextBox(ByVal value As Integer?, Optional leftPaddingZeroes As Byte = 0) As String
        If value.HasValue Then
            If leftPaddingZeroes = 0 Then
                Return FormatNumber(value.Value, 0)
            Else
                Return value.Value.ToString(StrDup(leftPaddingZeroes, "0"))
            End If
        Else
            Return String.Empty
        End If
    End Function

    Friend Function FromObjectIntegerToControlIntegerTextBox(ByVal ObjectValue As Integer?) As String
        If ObjectValue.HasValue Then
            Return FormatNumber(ObjectValue.Value, 0)
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectShortToControlTextBox(ByVal ObjectValue As Short?) As String
        If ObjectValue.HasValue Then
            Return FormatNumber(ObjectValue.Value, 0)
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectByteToControlTextBox(ByVal ObjectValue As Byte?) As String
        If ObjectValue.HasValue Then
            Return FormatNumber(ObjectValue.Value, 0)
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectByteToControlUpDown(ByVal ObjectValue As Byte?) As Decimal
        If ObjectValue.HasValue Then
            Return ObjectValue.Value
        Else
            Return Nothing
        End If
    End Function

    Friend Function FromObjectShortToControlUpDown(ByVal ObjectValue As Short?) As Decimal
        If ObjectValue.HasValue Then
            Return ObjectValue.Value
        Else
            Return Nothing
        End If
    End Function

    Friend Function FromObjectDecimalToControlTextBox(ByVal ObjectValue As Decimal?) As String
        If ObjectValue.HasValue Then
            Return ObjectValue.Value.ToString
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectDecimalToControlDoubleTextBox(ByVal ObjectValue As Decimal?) As String
        If ObjectValue.HasValue Then
            Return ObjectValue.Value.ToString
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectPercentToControlTextBox(ByVal ObjectValue As Decimal?) As String
        If ObjectValue.HasValue Then
            Return FormatNumber(ObjectValue.Value)
        Else
            Return ""
        End If
    End Function

    Friend Function FromObjectBooleanToControlCheckBox(ByVal ObjectValue As Boolean?) As CheckState
        If ObjectValue Is Nothing Then
            Return CheckState.Indeterminate
        ElseIf ObjectValue Then
            Return CheckState.Checked
        Else
            Return CheckState.Unchecked
        End If
    End Function

    Friend Function FromObjectDateToControlDateTimePicker(ByVal ObjectValue As Date?, Optional ByRef DateTimePickerControlToCheck As DateTimePicker = Nothing) As Date
        If DateTimePickerControlToCheck IsNot Nothing Then
            If ObjectValue Is Nothing Then
                DateTimePickerControlToCheck.Checked = False
            Else
                DateTimePickerControlToCheck.Checked = True
            End If
        End If

        If ObjectValue Is Nothing Then
            Return System.DateTime.Now
        Else
            Return System.Convert.ToDateTime(ObjectValue)
        End If
    End Function

    Friend Function FromObjectDateToControlDateTimePicker_OnlyDate(ByVal ObjectValue As Date?, Optional ByRef DateTimePickerControlToCheck As DateTimePicker = Nothing) As Date
        If DateTimePickerControlToCheck IsNot Nothing Then
            If ObjectValue Is Nothing Then
                DateTimePickerControlToCheck.Checked = False
            Else
                DateTimePickerControlToCheck.Checked = True
            End If
        End If

        If ObjectValue Is Nothing Then
            Return DateTime.Today
        Else
            If DateTimePickerControlToCheck IsNot Nothing Then
                If ObjectValue.Value < DateTimePickerControlToCheck.MinDate Then
                    If DateTimePickerControlToCheck.MinDate <= Date.Today Then
                        Return DateTime.Today
                    Else
                        Return DateTimePickerControlToCheck.MinDate
                    End If
                Else
                    Return ObjectValue.Value
                End If
            Else
                Return ObjectValue.Value
            End If
        End If
    End Function

    Friend Function FromObjectDateToControlDateTimePicker_OnlyTime(ByVal ObjectValue As Date?, Optional ByRef DateTimePickerControlToCheck As DateTimePicker = Nothing) As Date
        If DateTimePickerControlToCheck IsNot Nothing Then
            If ObjectValue Is Nothing Then
                DateTimePickerControlToCheck.Checked = False
            Else
                DateTimePickerControlToCheck.Checked = True
            End If
        End If

        If ObjectValue Is Nothing Then
            Return FromDateTimeToTime(DateTime.Now)
        Else
            Return FromDateTimeToTime(ObjectValue.Value)
        End If
    End Function

    Friend Function FromObjectTimeSpanToControlDateTimePicker(ByVal ObjectValue As TimeSpan?, Optional ByRef DateTimePickerControlToCheck As DateTimePicker = Nothing) As Date
        If DateTimePickerControlToCheck IsNot Nothing Then
            If ObjectValue Is Nothing Then
                DateTimePickerControlToCheck.Checked = False
            Else
                DateTimePickerControlToCheck.Checked = True
            End If
        End If

        If ObjectValue Is Nothing Then
            Return New Date(Now.Year, Now.Month, Now.Day, 0, 0, 0)
        Else
            Return New Date(Now.Year, Now.Month, Now.Day).AddHours(ObjectValue.Value.Hours).AddMinutes(ObjectValue.Value.Minutes)
        End If
    End Function

#End Region

#Region "De Controles a Objectos - TextBox (standard)"

    Friend Function FromControlTextBoxToObjectString(ByVal TextBoxText As String, Optional Conversion As VbStrConv = VbStrConv.None, Optional ByVal TrimText As Boolean = True) As String
        If TrimText Then
            TextBoxText = TextBoxText.Trim
        End If
        If TextBoxText = "" Then
            Return Nothing
        Else
            Return StrConv(TextBoxText, Conversion)
        End If
    End Function

    Friend Function FromControlTextBoxToObjectByte(ByVal TextBoxText As String) As Byte?
        Dim ConvertedValue As Byte

        Byte.TryParse(TextBoxText.Trim, Globalization.NumberStyles.None, My.Application.Culture, ConvertedValue)
        If ConvertedValue = 0 Then
            Return Nothing
        Else
            Return ConvertedValue
        End If
    End Function

    Friend Function FromControlTextBoxToObjectShort(ByVal TextBoxText As String) As Short?
        Dim ConvertedValue As Short

        If TextBoxText = "" Then
            Return Nothing
        Else
            Short.TryParse(TextBoxText.Trim, Globalization.NumberStyles.None, My.Application.Culture, ConvertedValue)
            Return ConvertedValue
        End If
    End Function

    Friend Function FromControlTextBoxToObjectInteger(ByVal TextBoxText As String) As Integer?
        Dim ConvertedValue As Integer

        If TextBoxText = "" Then
            Return Nothing
        Else
            Integer.TryParse(TextBoxText.Trim, Globalization.NumberStyles.AllowThousands, My.Application.Culture, ConvertedValue)
            Return ConvertedValue
        End If
    End Function

    Friend Function FromControlTextBoxToObjectDecimal(ByVal TextBoxText As String) As Decimal?
        Dim ConvertedValue As Decimal

        If TextBoxText = "" Then
            Return Nothing
        Else
            Decimal.TryParse(TextBoxText.Trim, Globalization.NumberStyles.Currency, My.Application.Culture, ConvertedValue)
            Return ConvertedValue
        End If
    End Function

#End Region

#Region "De Controles a Objectos - ComboBox"

    Friend Function FromControlComboBoxToObjectByte(ByVal ComboBoxSelectedValue As Object, Optional ValueForNull As Byte = CardonerSistemas.Constants.FIELD_VALUE_NOTSPECIFIED_BYTE) As Byte?
        If ComboBoxSelectedValue Is Nothing Then
            Return Nothing
        ElseIf CByte(ComboBoxSelectedValue) = ValueForNull Then
            Return Nothing
        Else
            Return CByte(ComboBoxSelectedValue)
        End If
    End Function

    Friend Function FromControlComboBoxToObjectShort(ByVal ComboBoxSelectedValue As Object, Optional ValueForNull As Short = CardonerSistemas.Constants.FIELD_VALUE_NOTSPECIFIED_SHORT) As Short?
        If ComboBoxSelectedValue Is Nothing Then
            Return Nothing
        ElseIf CShort(ComboBoxSelectedValue) = ValueForNull Then
            Return Nothing
        Else
            Return CShort(ComboBoxSelectedValue)
        End If
    End Function

    Friend Function FromControlComboBoxToObjectInteger(ByVal ComboBoxSelectedValue As Object, Optional ValueForNull As Integer = CardonerSistemas.Constants.FIELD_VALUE_NOTSPECIFIED_INTEGER) As Integer?
        If ComboBoxSelectedValue Is Nothing Then
            Return Nothing
        ElseIf CInt(ComboBoxSelectedValue) = ValueForNull Then
            Return Nothing
        Else
            Return CInt(ComboBoxSelectedValue)
        End If
    End Function

    Friend Function FromControlComboBoxToObjectString(ByVal ComboBoxSelectedValue As Object, Optional ValueForNull As String = Nothing) As String
        If ComboBoxSelectedValue Is Nothing Then
            Return Nothing
        ElseIf ComboBoxSelectedValue.ToString = ValueForNull Then
            Return Nothing
        Else
            Return ComboBoxSelectedValue.ToString
        End If
    End Function

#End Region

#Region "De Controles a Objectos - DateTimePicker"

    Friend Function FromControlDateTimePickerToObjectDate(ByVal DateTimePickerValue As Date, Optional DateTimePickerChecked As Boolean = True) As Date?
        If DateTimePickerChecked Then
            Return DateTimePickerValue
        Else
            Return Nothing
        End If
    End Function

    Friend Function FromControlTwoDateTimePickerToObjectDate(ByVal DateTimePickerValue_Date As Date, ByVal DateTimePickerValue_Time As Date, Optional DateTimePickerChecked As Boolean = True) As Date?
        If DateTimePickerChecked Then
            Return CDate(DateTimePickerValue_Date.ToShortDateString & " " & DateTimePickerValue_Time.ToShortTimeString)
        Else
            Return Nothing
        End If
    End Function

    Friend Function FromControlDateTimePickerToObjectTimeSpan(ByVal DateTimePickerValue As Date, Optional DateTimePickerChecked As Boolean = True) As TimeSpan?
        If DateTimePickerChecked Then
            Return New TimeSpan(DateTimePickerValue.Hour, DateTimePickerValue.Minute, 0)
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region "De Controles a Objectos - Misc"

    Friend Function FromControlUpDownToObjectByte(ByVal UpDownValue As Decimal, Optional valueForNothing As Decimal = 0) As Byte?
        If UpDownValue = valueForNothing Then
            Return Nothing
        Else
            Return CByte(UpDownValue)
        End If
    End Function

    Friend Function FromControlUpDownToObjectShort(ByVal UpDownValue As Decimal, Optional valueForNothing As Decimal = 0) As Short?
        If UpDownValue = valueForNothing Then
            Return Nothing
        Else
            Return CShort(UpDownValue)
        End If
    End Function

    Friend Function FromControlCheckBoxToObjectBoolean(ByVal CheckBoxCheckState As CheckState) As Boolean
        Select Case CheckBoxCheckState
            Case CheckState.Indeterminate
                Return Nothing
            Case CheckState.Checked
                Return True
            Case CheckState.Unchecked
                Return False
            Case Else
                Return Nothing
        End Select
    End Function

    Friend Function FromControlTagToObjectInteger(ByVal TagValue As Object) As Integer?
        If TypeOf (TagValue) Is Integer Then
            Return CInt(TagValue)
        Else
            Return Nothing
        End If
    End Function

#End Region

#Region "Validación"

    Friend Function ValidateCurrency(ByVal Value As String) As Boolean
        Dim ConvertedValue As Decimal

        Return Decimal.TryParse(Value, Globalization.NumberStyles.Currency, My.Application.Culture, ConvertedValue)
    End Function

    Friend Function ValidateDecimal(ByVal Value As String) As Boolean
        Dim ConvertedValue As Decimal

        Return Decimal.TryParse(Value, Globalization.NumberStyles.Float, My.Application.Culture, ConvertedValue)
    End Function

#End Region

#Region "Cambios de Formato"

    Friend Function DecimalToUString(ByVal value As Decimal) As String
        Return value.ToString.Replace(","c, "."c)
    End Function

    Friend Function UStringToDecimal(ByVal value As String) As Decimal
        Dim result As Decimal

        Decimal.TryParse(value.ToString.Replace("."c, ","c), result)
        Return result
    End Function

    Friend Function FromStringToPreventSQLInjection(ByVal Value As String) As String
        Return Value.Replace("'", "''")
    End Function

    Friend Function FromDateToUString(ByVal Value As Date) As String
        If Value = DateTime.MinValue Then
            Return Nothing
        Else
            Return Format(Value, "yyyyMMdd")
        End If
    End Function

    Friend Function FromStringToOnlyDigitsString(ByVal Value As String) As String
        Return Regex.Replace(Value, "\D", "")
    End Function

    Friend Function FromDateTimeToDate(ByVal Value As Date) As Date
        Return New Date(Value.Year, Value.Month, Value.Day)
    End Function

    Friend Function FromDateTimeToTime(ByVal Value As Date) As Date
        Return New Date(1, 1, 1, Value.Hour, Value.Minute, Value.Second, Value.Millisecond)
    End Function

    Friend Function FromDecimalToRoundedCurrency(ByVal Value As Decimal) As Decimal
        Dim nfi As New System.Globalization.NumberFormatInfo

        Return Math.Round(Value, nfi.CurrencyDecimalDigits)
    End Function

    Friend Function FromDoubleToRoundedCurrency(ByVal Value As Double) As Decimal
        Dim nfi As New System.Globalization.NumberFormatInfo

        Return Math.Round(CDec(Value), nfi.CurrencyDecimalDigits)
    End Function

#End Region

#Region "Recodificación de Strings"

    Friend Function FromOneEncodingToEncoding(ByRef SourceEncoding As Encoding, ByRef DestinationEncoding As Encoding, ByVal SourceString As String) As String
        ' Convert the string into a byte array.
        Dim SourceStringBytes As Byte() = SourceEncoding.GetBytes(SourceString)

        ' Perform the conversion from one encoding to the other.
        Dim DestinationStringBytes As Byte() = Encoding.Convert(SourceEncoding, DestinationEncoding, SourceStringBytes)

        ' Convert the new byte array into a char array and then into a string.
        Dim DestinationStringChars(DestinationEncoding.GetCharCount(DestinationStringBytes, 0, DestinationStringBytes.Length) - 1) As Char
        DestinationEncoding.GetChars(DestinationStringBytes, 0, DestinationStringBytes.Length, DestinationStringChars, 0)
        Dim DestinationString As New String(DestinationStringChars)

        Return DestinationString
    End Function

    Friend Function FromStringUnicodeToASCII(ByVal UnicodeString As String) As String
        ' Create two different encodings.
        Dim Unicode As Encoding = Encoding.Unicode
        Dim ASCII As Encoding = Encoding.ASCII

        Return FromOneEncodingToEncoding(Unicode, ASCII, UnicodeString)
    End Function

    Friend Function FromStringUnicodeToUTF8(ByVal UnicodeString As String) As String
        ' Create two different encodings.
        Dim Unicode As Encoding = Encoding.Unicode
        Dim UTF8 As Encoding = Encoding.UTF8

        Return FromOneEncodingToEncoding(Unicode, UTF8, UnicodeString)
    End Function

    Friend Function FromStringASCIIToUnicode(ByVal ASCIIString As String) As String
        ' Create two different encodings.
        Dim ASCII As Encoding = Encoding.ASCII
        Dim Unicode As Encoding = Encoding.Unicode

        Return FromOneEncodingToEncoding(ASCII, Unicode, ASCIIString)
    End Function

    Friend Function FromStringASCIIToUTF8(ByVal ASCIIString As String) As String
        ' Create two different encodings.
        Dim ASCII As Encoding = Encoding.ASCII
        Dim UTF8 As Encoding = Encoding.UTF8

        Return FromOneEncodingToEncoding(ASCII, UTF8, ASCIIString)
    End Function

    Friend Function FromStringUTF8ToASCII(ByVal UTF8String As String) As String
        ' Create two different encodings.
        Dim UTF8 As Encoding = Encoding.UTF8
        Dim ASCII As Encoding = Encoding.ASCII

        Return FromOneEncodingToEncoding(UTF8, ASCII, UTF8String)
    End Function

    Friend Function FromStringUTF8ToUnicode(ByVal UTF8String As String) As String
        ' Create two different encodings.
        Dim UTF8 As Encoding = Encoding.UTF8
        Dim Unicode As Encoding = Encoding.Unicode

        Return FromOneEncodingToEncoding(UTF8, Unicode, UTF8String)
    End Function

#End Region

#Region "Base de datos"

    Friend Function FromDBValueToDecimalValueOrZeroIfIsNull(ByRef value As Object) As Object
        If IsDBNull(value) Then
            Return 0
        Else
            Return Convert.ToDecimal(value)
        End If
    End Function

#End Region

End Module