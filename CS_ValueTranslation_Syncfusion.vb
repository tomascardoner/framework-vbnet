Imports Syncfusion.Windows.Forms.Tools

Module CS_ValueTranslation_Syncfusion

#Region "De Objectos a Controles - TextBox (SyncFusion)"

    ' Integer as Byte
    Friend Sub FromValueToControl(ByVal value As Byte?, ByRef control As IntegerTextBox)
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

    ' Integer as Short
    Friend Sub FromValueToControl(ByVal value As Short?, ByRef control As IntegerTextBox)
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

    ' Double
    Friend Sub FromValueToControl(ByVal value As Decimal?, ByRef control As DoubleTextBox)
        If value.HasValue Then
            control.DoubleValue = value.Value
        Else
            If control.AllowNull Then
                control.BindableValue = Nothing
            Else
                control.DoubleValue = control.MinValue
            End If
        End If
    End Sub

    ' Currency
    Friend Sub FromValueToControl(ByVal value As Decimal?, ByRef control As CurrencyTextBox)
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

    ' Percent
    Friend Sub FromValueToControl(ByVal value As Decimal?, ByRef control As PercentTextBox)
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

    ' Byte
    Friend Function FromControlToByte(ByRef control As IntegerTextBox) As Byte?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToByte(control.IntegerValue)
        End If
    End Function

    ' Short
    Friend Function FromControlToShort(ByRef control As IntegerTextBox) As Short?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToInt16(control.IntegerValue)
        End If
    End Function

    ' Integer
    Friend Function FromControlToInteger(ByRef control As IntegerTextBox) As Integer?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToInt32(control.IntegerValue)
        End If
    End Function

    ' Decimal
    Friend Function FromControlToDecimal(ByVal control As DoubleTextBox) As Decimal?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToDecimal(control.DoubleValue)
        End If
    End Function

    Friend Function FromControlToDecimal(ByRef control As CurrencyTextBox) As Decimal?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return control.DecimalValue
        End If
    End Function

    Friend Function FromControlToDecimal(ByRef control As PercentTextBox) As Decimal?
        If control.AllowNull AndAlso control.IsNull Then
            Return Nothing
        Else
            Return Convert.ToDecimal(control.DoubleValue)
        End If
    End Function

#End Region

End Module