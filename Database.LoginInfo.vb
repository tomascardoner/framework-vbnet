Namespace CardonerSistemas.Database

    Public Class LoginInfo

        Private Sub Me_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
            If e.KeyChar = ChrW(Keys.Return) Then
                buttonAceptar_Click()
            ElseIf e.KeyChar = ChrW(Keys.Escape) Then
                buttonCancelar_Click()
            End If
        End Sub

        Private Sub buttonAceptar_Click() Handles buttonAceptar.Click
            If String.IsNullOrWhiteSpace(textboxUsuario.Text) Then
                MessageBox.Show("Debe ingresar el usuario.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Information)
                textboxUsuario.Focus()
                Exit Sub
            End If

            Me.DialogResult = Windows.Forms.DialogResult.OK
        End Sub

        Private Sub buttonCancelar_Click() Handles buttonCancelar.Click
            Me.DialogResult = Windows.Forms.DialogResult.Cancel
        End Sub

    End Class
    
End Namespace