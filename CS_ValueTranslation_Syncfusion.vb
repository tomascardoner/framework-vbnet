Imports System.Text
Imports System.Text.RegularExpressions

Module CS_ValueTranslation_Syncfusion

#Region "De Objectos a Controles - TextBox (SyncFusion)"

    Friend Sub FromValueByteToControlIntegerTextBox(ByVal value As Byte?, ByRef control As Syncfusion.Windows.Forms.Tools.IntegerTextBox)
        If value.HasValue Then
            control.IntegerValue = value.Value
        Else
            If control.AllowNull Then
                control.BindableValue = Nothing
            Else
                control.IntegerValue = control.MinValue
            End If
        End If
    End Sub

    Friend Sub FromValueDecimalToControlCurrencyTextBox(ByVal value As Decimal?, ByRef control As Syncfusion.Windows.Forms.Tools.CurrencyTextBox)
        If value.HasValue Then
            control.DecimalValue = value.Value
        Else
            If control.AllowNull Then
                control.BindableValue = Nothing
            Else
                control.DecimalValue = control.MinValue
            End If
        End If
    End Sub

    Friend Sub FromValueDecimalToControlPercentTextBox(ByVal value As Decimal?, ByRef control As Syncfusion.Windows.Forms.Tools.PercentTextBox)
        If value.HasValue Then
            control.PercentValue = value.Value
        Else
            If control.AllowNull Then
                control.BindableValue = Nothing
            Else
                control.PercentValue = control.MinValue
            End If
        End If
    End Sub

#End Region

#Region "De Controles a Objectos - TextBox (SyncFusion)"

    Friend Function FromControlIntegerTextBoxToValueByte(ByRef control As Syncfusion.Windows.Forms.Tools.IntegerTextBox) As Byte?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToByte(control.IntegerValue)
        End If
    End Function

    Friend Function FromControlIntegerTextBoxToValueShort(ByRef control As Syncfusion.Windows.Forms.Tools.IntegerTextBox) As Short?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToInt16(control.IntegerValue)
        End If
    End Function

    Friend Function FromControlIntegerTextBoxToValueInteger(ByRef control As Syncfusion.Windows.Forms.Tools.IntegerTextBox) As Integer?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToInt32(control.IntegerValue)
        End If
    End Function

    Friend Function FromControlDoubleTextBoxToObjectDouble(ByVal BindableValue As Object) As Double?
        If BindableValue Is Nothing Or Not IsNumeric(BindableValue) Then
            Return Nothing
        Else
            Return Convert.ToDouble(BindableValue)
        End If
    End Function

    Friend Function FromControlDoubleTextBoxToObjectDecimal(ByVal BindableValue As Object) As Decimal?
        If BindableValue Is Nothing Or Not IsNumeric(BindableValue) Then
            Return Nothing
        Else
            Return Convert.ToDecimal(BindableValue)
        End If
    End Function

    Friend Function FromControlCurrencyTextBoxToObjectDecimal(ByRef control As Syncfusion.Windows.Forms.Tools.CurrencyTextBox) As Decimal?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return control.DecimalValue
        End If
    End Function

    Friend Function FromControlPercentTextBoxToObjectDecimal(ByRef control As Syncfusion.Windows.Forms.Tools.PercentTextBox) As Decimal?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToDecimal(control.PercentValue)
        End If
    End Function

#End Region

End Module