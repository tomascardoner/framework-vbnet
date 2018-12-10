Public Class CS_Control_TextBox_Currency
    ' How to use:
    ' Using as a control:
    '   1. When you want to change the value of the text programmatically follow this statement:
    '           CurrencyTextBox1.Text = "9876543.21" or CurrencyTextBox1.Text = MyDecimal
    '      In that case "9876543.21" can be string, or MyDecimal can be both string and any numeric
    '      variable. 
    '   2. If you want to get the value of the currency not in the formatted mode, you can use
    '           MyDecimal = CurrencyTextBox1.Value ' Dont use .Text since it will output the formatted code.

    Inherits System.Windows.Forms.TextBox

    Private mDecimalPeriodPressed As Boolean = False
    Private mDecimalPeriod As Boolean = False
    Private mNumberFormat As System.Globalization.NumberFormatInfo = My.Application.Culture.NumberFormat
    Private mBuffer As String = ""

    'Public Overrides Property Text As String
    '    Get
    '        Return MyBase.Text
    '    End Get
    '    Set(ByVal value As String)
    '        Try
    '            MyBase.Text = FormatCurrency(Decimal.Parse(value, Globalization.NumberStyles.Currency, mNumberFormat).ToString)
    '        Catch ex As Exception
    '            MyBase.Text = "0"
    '        End Try
    '    End Set
    'End Property

    Public ReadOnly Property Value As Decimal?
        Get
            If MyBase.Text.Length = 0 Then
                Return Nothing
            Else
                Return Decimal.Parse(MyBase.Text, Globalization.NumberStyles.Currency, mNumberFormat)
            End If
        End Get
    End Property

    Private Sub Me_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.GotFocus
        If Not MyBase.ReadOnly Then
            If MyBase.Text.Length = 0 Then
                MyBase.Text = "0"
            End If
            MyBase.Text = Decimal.Parse(MyBase.Text, Globalization.NumberStyles.Currency, mNumberFormat).ToString
        End If
    End Sub

    Private Sub Me_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Decimal Then
            mDecimalPeriodPressed = True
        End If
    End Sub

    Private Sub Me_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        mBuffer = MyBase.Text
        If e.KeyChar = Chr(Keys.Back) Then
            ' Se presionó Backspace
            If MyBase.Text.Length > 0 Then
                Exit Sub
            End If
            mDecimalPeriod = False
            Exit Sub
        ElseIf Char.IsDigit(e.KeyChar) = False Then
            ' No es un número, así que verifico que sea un caracter válido

            If mDecimalPeriodPressed Then
                ' Se presionó la tecla decimal del teclado numérico, así que lo convierto en el separador decimal del sistema
                e.KeyChar = mNumberFormat.CurrencyDecimalSeparator.Chars(0)
                mDecimalPeriodPressed = False
            End If

            If e.KeyChar = mNumberFormat.CurrencyDecimalSeparator.Chars(0) Then
                ' Si es el separador decimal, verifico que no exista
                If MyBase.Text.Length = 0 And MyBase.SelectionLength = 0 Then
                    e.Handled = True
                    Exit Sub
                End If
                If mDecimalPeriod = False And MyBase.SelectionLength = 0 Then
                    mDecimalPeriod = True
                    Exit Sub
                Else
                    e.Handled = True
                End If
            End If
            e.Handled = True
            Exit Sub
        End If
        If IsNumber(Text) = False Then
            MyBase.Text = mBuffer
        End If
    End Sub

    Private Sub Me_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
        MyBase.SelectAll()
    End Sub

    Private Sub Me_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        'If MyBase.Enabled = False Or MyBase.ReadOnly = True Then
        '    MyBase.Text = Decimal.Parse(MyBase.Text, Globalization.NumberStyles.Currency, mNumberFormat).ToString(CURRFORM, mNumberFormat)
        'End If
    End Sub

    Private Sub Me_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.LostFocus
        If Not MyBase.ReadOnly Then
            If MyBase.Text.Length = 0 Then
                MyBase.Text = "0"
            End If
            MyBase.Text = FormatCurrency(Decimal.Parse(MyBase.Text, Globalization.NumberStyles.Currency, mNumberFormat).ToString)
            mDecimalPeriod = False
        End If
    End Sub

    Private Function IsNumber(ByVal str As String) As Boolean
        Dim iCounter As Integer
        Dim bValidator As Boolean
        For iCounter = 0 To str.Length - 1
            If Char.IsDigit(str.Chars(iCounter)) = False Then
                If str.Chars(iCounter) = mNumberFormat.CurrencyDecimalSeparator.Chars(0) And bValidator = False Then
                    bValidator = True
                Else
                    Return False
                    Exit Function
                End If
            End If
        Next

        Return True
    End Function
End Class