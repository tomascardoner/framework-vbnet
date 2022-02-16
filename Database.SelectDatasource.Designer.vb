Namespace CardonerSistemas.Database
    <Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
    Partial Class SelectDatasource
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
            Me.comboboxDataSource = New System.Windows.Forms.ComboBox()
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
            Me.toolstripMain.TabIndex = 5
            '
            'buttonCancelar
            '
            Me.buttonCancelar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
            Me.buttonCancelar.Image = My.Resources.ImageCancelar32
            Me.buttonCancelar.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            Me.buttonCancelar.ImageTransparentColor = System.Drawing.Color.Magenta
            Me.buttonCancelar.Name = "buttonCancelar"
            Me.buttonCancelar.Size = New System.Drawing.Size(89, 36)
            Me.buttonCancelar.Text = "Cancelar"
            '
            'buttonAceptar
            '
            Me.buttonAceptar.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
            Me.buttonAceptar.Image = My.Resources.ImageAceptar32
            Me.buttonAceptar.ImageScaling = System.Windows.Forms.ToolStripItemImageScaling.None
            Me.buttonAceptar.ImageTransparentColor = System.Drawing.Color.Magenta
            Me.buttonAceptar.Name = "buttonAceptar"
            Me.buttonAceptar.Size = New System.Drawing.Size(84, 36)
            Me.buttonAceptar.Text = "Aceptar"
            '
            'pictureboxMain
            '
            Me.pictureboxMain.Image = My.Resources.ImageDatabaseSelect48
            Me.pictureboxMain.Location = New System.Drawing.Point(12, 42)
            Me.pictureboxMain.Name = "pictureboxMain"
            Me.pictureboxMain.Size = New System.Drawing.Size(48, 48)
            Me.pictureboxMain.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
            Me.pictureboxMain.TabIndex = 6
            Me.pictureboxMain.TabStop = False
            '
            'comboboxDataSource
            '
            Me.comboboxDataSource.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            Me.comboboxDataSource.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            Me.comboboxDataSource.FormattingEnabled = True
            Me.comboboxDataSource.Location = New System.Drawing.Point(66, 57)
            Me.comboboxDataSource.Name = "comboboxDataSource"
            Me.comboboxDataSource.Size = New System.Drawing.Size(285, 24)
            Me.comboboxDataSource.TabIndex = 7
            '
            'CS_Database_SelectSource
            '
            Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            Me.ClientSize = New System.Drawing.Size(363, 105)
            Me.ControlBox = False
            Me.Controls.Add(Me.comboboxDataSource)
            Me.Controls.Add(Me.pictureboxMain)
            Me.Controls.Add(Me.toolstripMain)
            Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
            Me.KeyPreview = True
            Me.Name = "CS_Database_SelectSource"
            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Text = "Seleccione el origen de los datos"
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
        Friend WithEvents comboboxDataSource As System.Windows.Forms.ComboBox
    End Class
End Namespace