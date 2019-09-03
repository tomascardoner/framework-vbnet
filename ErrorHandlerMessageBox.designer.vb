Namespace CardonerSistemas

    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class ErrorHandlerMessageBox
        Inherits System.Windows.Forms.Form

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(ErrorHandlerMessageBox))
            Me.picError = New System.Windows.Forms.PictureBox()
            Me.chkDetail = New System.Windows.Forms.CheckBox()
            Me.btnClose = New System.Windows.Forms.Button()
            Me.lblFriendlyMessage = New System.Windows.Forms.Label()
            Me.txtMessageData = New System.Windows.Forms.TextBox()
            Me.txtStackTraceData = New System.Windows.Forms.TextBox()
            Me.lblStackTrace = New System.Windows.Forms.Label()
            Me.lblMessage = New System.Windows.Forms.Label()
            Me.lblSourceData = New System.Windows.Forms.Label()
            Me.lblSource = New System.Windows.Forms.Label()
            CType(Me.picError, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'picError
            '
            Me.picError.Image = CType(resources.GetObject("picError.Image"), System.Drawing.Image)
            Me.picError.Location = New System.Drawing.Point(16, 34)
            Me.picError.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.picError.Name = "picError"
            Me.picError.Size = New System.Drawing.Size(48, 48)
            Me.picError.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
            Me.picError.TabIndex = 15
            Me.picError.TabStop = False
            '
            'chkDetail
            '
            Me.chkDetail.Appearance = System.Windows.Forms.Appearance.Button
            Me.chkDetail.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.chkDetail.Location = New System.Drawing.Point(529, 132)
            Me.chkDetail.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.chkDetail.Name = "chkDetail"
            Me.chkDetail.Size = New System.Drawing.Size(101, 30)
            Me.chkDetail.TabIndex = 14
            Me.chkDetail.Text = "&Detalle >>"
            Me.chkDetail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
            '
            'btnClose
            '
            Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.btnClose.Location = New System.Drawing.Point(417, 132)
            Me.btnClose.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.btnClose.Name = "btnClose"
            Me.btnClose.Size = New System.Drawing.Size(101, 30)
            Me.btnClose.TabIndex = 13
            Me.btnClose.Text = "&Cerrar"
            '
            'lblFriendlyMessage
            '
            Me.lblFriendlyMessage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblFriendlyMessage.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.lblFriendlyMessage.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblFriendlyMessage.Location = New System.Drawing.Point(95, 11)
            Me.lblFriendlyMessage.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblFriendlyMessage.Name = "lblFriendlyMessage"
            Me.lblFriendlyMessage.Size = New System.Drawing.Size(535, 105)
            Me.lblFriendlyMessage.TabIndex = 12
            '
            'txtMessageData
            '
            Me.txtMessageData.BackColor = System.Drawing.SystemColors.Control
            Me.txtMessageData.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.txtMessageData.Location = New System.Drawing.Point(119, 327)
            Me.txtMessageData.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.txtMessageData.MaxLength = 0
            Me.txtMessageData.Multiline = True
            Me.txtMessageData.Name = "txtMessageData"
            Me.txtMessageData.ReadOnly = True
            Me.txtMessageData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            Me.txtMessageData.Size = New System.Drawing.Size(511, 106)
            Me.txtMessageData.TabIndex = 21
            '
            'txtStackTraceData
            '
            Me.txtStackTraceData.BackColor = System.Drawing.SystemColors.Control
            Me.txtStackTraceData.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.txtStackTraceData.Location = New System.Drawing.Point(119, 208)
            Me.txtStackTraceData.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.txtStackTraceData.MaxLength = 0
            Me.txtStackTraceData.Multiline = True
            Me.txtStackTraceData.Name = "txtStackTraceData"
            Me.txtStackTraceData.ReadOnly = True
            Me.txtStackTraceData.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            Me.txtStackTraceData.Size = New System.Drawing.Size(511, 99)
            Me.txtStackTraceData.TabIndex = 19
            '
            'lblStackTrace
            '
            Me.lblStackTrace.AutoSize = True
            Me.lblStackTrace.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.lblStackTrace.Location = New System.Drawing.Point(17, 208)
            Me.lblStackTrace.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblStackTrace.Name = "lblStackTrace"
            Me.lblStackTrace.Size = New System.Drawing.Size(55, 17)
            Me.lblStackTrace.TabIndex = 18
            Me.lblStackTrace.Text = "Origen:"
            '
            'lblMessage
            '
            Me.lblMessage.AutoSize = True
            Me.lblMessage.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.lblMessage.Location = New System.Drawing.Point(16, 331)
            Me.lblMessage.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblMessage.Name = "lblMessage"
            Me.lblMessage.Size = New System.Drawing.Size(86, 17)
            Me.lblMessage.TabIndex = 20
            Me.lblMessage.Text = "Descripción:"
            '
            'lblSourceData
            '
            Me.lblSourceData.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
            Me.lblSourceData.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.lblSourceData.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.lblSourceData.Location = New System.Drawing.Point(119, 174)
            Me.lblSourceData.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblSourceData.Name = "lblSourceData"
            Me.lblSourceData.Size = New System.Drawing.Size(512, 21)
            Me.lblSourceData.TabIndex = 17
            '
            'lblSource
            '
            Me.lblSource.AutoSize = True
            Me.lblSource.FlatStyle = System.Windows.Forms.FlatStyle.System
            Me.lblSource.Location = New System.Drawing.Point(17, 174)
            Me.lblSource.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
            Me.lblSource.Name = "lblSource"
            Me.lblSource.Size = New System.Drawing.Size(67, 17)
            Me.lblSource.TabIndex = 16
            Me.lblSource.Text = "Contexto:"
            '
            'CardonerSistemas.ErrorHandlerMessageBox
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(643, 444)
            Me.Controls.Add(Me.txtMessageData)
            Me.Controls.Add(Me.txtStackTraceData)
            Me.Controls.Add(Me.lblStackTrace)
            Me.Controls.Add(Me.lblMessage)
            Me.Controls.Add(Me.lblSourceData)
            Me.Controls.Add(Me.lblSource)
            Me.Controls.Add(Me.picError)
            Me.Controls.Add(Me.chkDetail)
            Me.Controls.Add(Me.btnClose)
            Me.Controls.Add(Me.lblFriendlyMessage)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
            Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
            Me.MaximizeBox = False
            Me.MinimizeBox = False
            Me.Name = "CardonerSistemas.ErrorHandlerMessageBox"
            Me.ShowIcon = False
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            Me.Text = "CardonerSistemas.ErrorHandlerMessageBox"
            Me.TopMost = True
            CType(Me.picError, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents picError As System.Windows.Forms.PictureBox
        Friend WithEvents chkDetail As System.Windows.Forms.CheckBox
        Friend WithEvents btnClose As System.Windows.Forms.Button
        Friend WithEvents lblFriendlyMessage As System.Windows.Forms.Label
        Friend WithEvents txtMessageData As System.Windows.Forms.TextBox
        Friend WithEvents txtStackTraceData As System.Windows.Forms.TextBox
        Friend WithEvents lblStackTrace As System.Windows.Forms.Label
        Friend WithEvents lblMessage As System.Windows.Forms.Label
        Friend WithEvents lblSourceData As System.Windows.Forms.Label
        Friend WithEvents lblSource As System.Windows.Forms.Label
    End Class

End Namespace