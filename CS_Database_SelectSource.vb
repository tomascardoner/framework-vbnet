Public Class CS_Database_SelectSource

    Private Sub Me_Load() Handles Me.Load

    End Sub

    Private Sub Me_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = ChrW(Keys.Return) Then
            buttonAceptar_Click()
        ElseIf e.KeyChar = ChrW(Keys.Escape) Then
            buttonCancelar_Click()
        End If
    End Sub

    Private Sub Me_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
    End Sub

    Private Sub buttonAceptar_Click() Handles buttonAceptar.Click
        If comboboxDataSource.SelectedIndex = -1 Then
            MsgBox("Debe seleccionar el origen de los datos.", vbInformation, My.Application.Info.Title)
            comboboxDataSource.Focus()
            Exit Sub
        End If

        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub buttonCancelar_Click() Handles buttonCancelar.Click
        Me.DialogResult = Windows.Forms.DialogResult.Cancel
    End Sub

End Class