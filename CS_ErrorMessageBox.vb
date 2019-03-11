Public Class CS_ErrorMessageBox

    Private Sub chkDetail_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkDetail.CheckedChanged
        If chkDetail.Checked() Then
            Me.Height = txtMessageData.Location.Y + txtMessageData.Height + 50
        Else
            Me.Height = chkDetail.Location.Y + chkDetail.Height + 50
        End If
        Me.CenterToScreen()
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub frmErrorMessageBox_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Text = My.Application.Info.Title
        chkDetail_CheckedChanged(chkDetail, New System.EventArgs())
    End Sub

End Class