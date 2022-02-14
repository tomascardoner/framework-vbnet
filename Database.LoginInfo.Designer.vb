Namespace CardonerSistemas.Database
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class LoginInfo
        Inherits System.Windows.Forms.Form

        'Form overrides dispose to clean up the component list.
        <System.Diagnostics.DebuggerNonUserCode()>
        Protected Overrides Sub Dispose(ByVal disposing As Boolean)
            Try
                If disposing AndAlso components IsNot Nothing Then
                    components.Dispose()
                End If
            Finally
                MyBase.Dispose(disposing)
            End Try
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Me.toolstripMain = New System.Windows.Forms.ToolStrip()
            Me.buttonCancelar = New System.Windows.Forms.ToolStripButton()
            Me.buttonAceptar = New System.Windows.Forms.ToolStripButton()
            Me.pictureboxMain = New System.Windows.Forms.PictureBox()
            Me.textboxPassword = New System.Windows.Forms.TextBox()
            Me.textboxUsuario = New System.Windows.Forms.TextBox()
            Me.labelPassword = New System.Windows.Forms.Label()
            Me.labelUsuario = New System.Windows.Forms.Label()
            Me.toolstripMain.SuspendLayout()
            CType(Me.pictureboxMain, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.SuspendLayout()
            '
            'toolstripMain
            '
            Me.toolstripMain.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
            Me.toolstripMain.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.buttonCancelar, Me.buttonAceptar})
            Me.toolstripMain.Location = New System.Drawing.Point(0, 0)
            Me.toolstripMain.Name = "toolstripMain"
            Me.toolstripMain.Size = New System.Drawing.Size(363, 39)
            Me.toolstripMain.TabIndex = 4
            '
            'buttonCancelar
            '
            Me.buttonCancelar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
            Me.buttonCancelar.Image = Global.CSBomberos.My.Resources.Resources.ImageCancelar32
            Me.buttonCancelar.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            Me.buttonCancelar.ImageTransparentColor = System.Drawing.Color.Magenta
            Me.buttonCancelar.Name = "buttonCancelar"
            Me.buttonCancelar.Size = New System.Drawing.Size(89, 36)
            Me.buttonCancelar.Text = "Cancelar"
            '
            'buttonAceptar
            '
            Me.buttonAceptar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
            Me.buttonAceptar.Image = Global.CSBomberos.My.Resources.Resources.ImageAceptar32
            Me.buttonAceptar.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            Me.buttonAceptar.ImageTransparentColor = System.Drawing.Color.Magenta
            Me.buttonAceptar.Name = "buttonAceptar"
            Me.buttonAceptar.Size = New System.Drawing.Size(84, 36)
            Me.buttonAceptar.Text = "Aceptar"
            '
            'pictureboxMain
            '
            Me.pictureboxMain.Image = Global.CSBomberos.My.Resources.Resources.ImageDatabaseLoginInfo48
            Me.pictureboxMain.Location = New System.Drawing.Point(12, 42)
            Me.pictureboxMain.Name = "pictureboxMain"
            Me.pictureboxMain.Size = New System.Drawing.Size(48, 48)
            Me.pictureboxMain.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
            Me.pictureboxMain.TabIndex = 6
            Me.pictureboxMain.TabStop = False
            '
            'textboxPassword
            '
            Me.textboxPassword.Location = New System.Drawing.Point(136, 73)
            Me.textboxPassword.MaxLength = 128
            Me.textboxPassword.Name = "textboxPassword"
            Me.textboxPassword.Size = New System.Drawing.Size(215, 20)
            Me.textboxPassword.TabIndex = 3
            Me.textboxPassword.UseSystemPasswordChar = True
            '
            'textboxUsuario
            '
            Me.textboxUsuario.Location = New System.Drawing.Point(136, 43)
            Me.textboxUsuario.MaxLength = 20
            Me.textboxUsuario.Name = "textboxUsuario"
            Me.textboxUsuario.Size = New System.Drawing.Size(215, 20)
            Me.textboxUsuario.TabIndex = 1
            '
            'labelPassword
            '
            Me.labelPassword.AutoSize = True
            Me.labelPassword.Location = New System.Drawing.Point(66, 76)
            Me.labelPassword.Name = "labelPassword"
            Me.labelPassword.Size = New System.Drawing.Size(64, 13)
            Me.labelPassword.TabIndex = 2
            Me.labelPassword.Text = "Contraseña:"
            '
            'labelUsuario
            '
            Me.labelUsuario.AutoSize = True
            Me.labelUsuario.Location = New System.Drawing.Point(66, 46)
            Me.labelUsuario.Name = "labelUsuario"
            Me.labelUsuario.Size = New System.Drawing.Size(46, 13)
            Me.labelUsuario.TabIndex = 0
            Me.labelUsuario.Text = "Usuario:"
            '
            'LoginInfo
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(363, 105)
            Me.ControlBox = False
            Me.Controls.Add(Me.textboxPassword)
            Me.Controls.Add(Me.textboxUsuario)
            Me.Controls.Add(Me.labelPassword)
            Me.Controls.Add(Me.labelUsuario)
            Me.Controls.Add(Me.pictureboxMain)
            Me.Controls.Add(Me.toolstripMain)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.KeyPreview = True
            Me.Name = "LoginInfo"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Ingrese la información de inicio de sesión a la base de datos"
            Me.TopMost = True
            Me.toolstripMain.ResumeLayout(False)
            Me.toolstripMain.PerformLayout()
            CType(Me.pictureboxMain, System.ComponentModel.ISupportInitialize).EndInit()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub
        Friend WithEvents toolstripMain As System.Windows.Forms.ToolStrip
        Friend WithEvents buttonCancelar As System.Windows.Forms.ToolStripButton
        Friend WithEvents buttonAceptar As System.Windows.Forms.ToolStripButton
        Friend WithEvents pictureboxMain As System.Windows.Forms.PictureBox
        Friend WithEvents textboxPassword As TextBox
        Friend WithEvents textboxUsuario As TextBox
        Friend WithEvents labelPassword As Label
        Friend WithEvents labelUsuario As Label
    End Class
End Namespace