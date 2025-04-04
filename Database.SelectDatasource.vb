Namespace CardonerSistemas.Database

    Public Class SelectDatasource

        Private Sub Me_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
            If e.KeyChar = ChrW(Keys.Return) Then
                Aceptar_Click()
            ElseIf e.KeyChar = ChrW(Keys.Escape) Then
                Cancelar_Click()
            End If
        End Sub

        Private Sub Aceptar_Click() Handles buttonAceptar.Click
            If comboboxDataSource.SelectedIndex = -1 Then
                MsgBox("Debe seleccionar el origen de los datos.", vbInformation, My.Application.Info.Title)
                comboboxDataSource.Focus()
                Return
            End If

            Me.DialogResult = Windows.Forms.DialogResult.OK
        End Sub

        Private Sub Cancelar_Click() Handles buttonCancelar.Click
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
        End Sub

    End Class
    
End Namespace