Namespace CardonerSistemas.Controls

    Module ComboBox

        Friend Enum SelectedItemOptions
            None
            First
            NoneOrFirstIfUnique
            Last
            Current
            CurrentOrFirst
            CurrentOrFirstIfUnique
            CurrentOrLast
            Value
            ValueOrFirst
            ValueOrFirstIfUnique
            ValueOrLast
        End Enum

        Friend Sub SetSelectedValue(ByRef ComboBoxControl As System.Windows.Forms.ComboBox, Optional SelectedItemOption As SelectedItemOptions = SelectedItemOptions.None, Optional ValueToSelect As Object = Nothing, Optional ValueForNull As Object = Nothing)
            Dim SelectedValue As Object

            If ComboBoxControl.Items.Count > 0 Then
                SelectedValue = ComboBoxControl.SelectedValue
                If ValueToSelect Is Nothing Then
                    ValueToSelect = ValueForNull
                End If

                Select Case SelectedItemOption
                    Case SelectedItemOptions.None
                        ComboBoxControl.SelectedIndex = -1

                    Case SelectedItemOptions.First
                        ComboBoxControl.SelectedIndex = 0

                    Case SelectedItemOptions.NoneOrFirstIfUnique
                        If ComboBoxControl.Items.Count = 1 Then
                            ComboBoxControl.SelectedIndex = 0
                        Else
                            ComboBoxControl.SelectedIndex = -1
                        End If

                    Case SelectedItemOptions.Last
                        ComboBoxControl.SelectedIndex = ComboBoxControl.Items.Count - 1

                    Case SelectedItemOptions.Current
                        ComboBoxControl.SelectedValue = SelectedValue

                    Case SelectedItemOptions.CurrentOrFirst
                        ComboBoxControl.SelectedValue = SelectedValue
                        If ComboBoxControl.SelectedValue Is Nothing Then
                            ComboBoxControl.SelectedIndex = 0
                        End If

                    Case SelectedItemOptions.CurrentOrFirstIfUnique
                        ComboBoxControl.SelectedValue = SelectedValue
                        If ComboBoxControl.SelectedValue Is Nothing Then
                            If ComboBoxControl.Items.Count = 1 Then
                                ComboBoxControl.SelectedIndex = 0
                            Else
                                ComboBoxControl.SelectedIndex = -1
                            End If
                        End If

                    Case SelectedItemOptions.CurrentOrLast
                        ComboBoxControl.SelectedValue = SelectedValue
                        If ComboBoxControl.SelectedValue Is Nothing Then
                            ComboBoxControl.SelectedIndex = ComboBoxControl.Items.Count - 1
                        End If

                    Case SelectedItemOptions.Value
                        If ValueToSelect Is Nothing Then
                            ComboBoxControl.SelectedIndex = -1
                        Else
                            ComboBoxControl.SelectedValue = ValueToSelect
                        End If

                    Case SelectedItemOptions.ValueOrFirst
                        If ValueToSelect Is Nothing Then
                            ComboBoxControl.SelectedIndex = 0
                        Else
                            ComboBoxControl.SelectedValue = ValueToSelect
                            If ComboBoxControl.SelectedValue Is Nothing Then
                                ComboBoxControl.SelectedIndex = 0
                            End If
                        End If

                    Case SelectedItemOptions.ValueOrFirstIfUnique
                        If ValueToSelect Is Nothing Then
                            If ComboBoxControl.Items.Count = 1 Then
                                ComboBoxControl.SelectedIndex = 0
                            Else
                                ComboBoxControl.SelectedIndex = -1
                            End If
                        Else
                            ComboBoxControl.SelectedValue = ValueToSelect
                            If ComboBoxControl.SelectedValue Is Nothing Then
                                If ComboBoxControl.Items.Count = 1 Then
                                    ComboBoxControl.SelectedIndex = 0
                                Else
                                    ComboBoxControl.SelectedItem = -1
                                End If
                            End If
                        End If

                    Case SelectedItemOptions.ValueOrLast
                        If ValueToSelect Is Nothing Then
                            If ComboBoxControl.SelectedValue Is Nothing Then
                                ComboBoxControl.SelectedIndex = ComboBoxControl.Items.Count - 1
                            End If
                        Else
                            ComboBoxControl.SelectedValue = ValueToSelect
                            If ComboBoxControl.SelectedValue Is Nothing Then
                                ComboBoxControl.SelectedIndex = ComboBoxControl.Items.Count - 1
                            End If
                        End If
                End Select
            End If
        End Sub

        Friend Sub SetSelectedIndexByDisplayValue(ByRef ComboBoxControl As System.Windows.Forms.ComboBox, ByVal DisplayValueToSelect As String)
            ComboBoxControl.Text = DisplayValueToSelect
        End Sub

    End Module

End Namespace