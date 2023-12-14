Namespace CardonerSistemas
    Module Forms

        Friend Sub MdiChildShow(ByRef mdiForm As Form, ByRef childForm As Form, ByVal centerForm As Boolean)
            mdiForm.Cursor = Cursors.WaitCursor

            childForm.MdiParent = mdiForm
            If centerForm Then
                CenterToParent(mdiForm, childForm)
            Else
                MdiChildPositionAndSizeToFit(mdiForm, childForm)
            End If
            childForm.Show()
            If childForm.WindowState = FormWindowState.Minimized Then
                childForm.WindowState = FormWindowState.Normal
            End If
            childForm.Focus()

            mdiForm.Cursor = Cursors.Default
        End Sub

        Friend Sub MdiChildPositionAndSizeToFit(ByRef mdiForm As Form, ByRef childForm As Form)
            With childForm
                .SuspendLayout()

                .MdiParent = mdiForm

                If .WindowState <> FormWindowState.Normal Then
                    .WindowState = FormWindowState.Normal
                End If

                .Dock = DockStyle.Fill

                .ResumeLayout(True)
            End With
        End Sub

        Friend Sub MdiChildCenterToClientArea(ByRef childForm As Form, ByVal clientSize As Drawing.Size)
            With childForm
                If .WindowState <> FormWindowState.Normal Then
                    .WindowState = FormWindowState.Normal
                End If

                .Top = CInt((clientSize.Height - .Height) / 2)
                .Left = CInt((clientSize.Width - .Width) / 2)
            End With
        End Sub

        Friend Function MdiChildIsLoaded(ByRef mdiForm As Form, ByVal formName As String) As Boolean
            For Each ChildForm As Form In mdiForm.MdiChildren
                If ChildForm.Name.ToLower() = formName.ToLower() Then
                    Return True
                End If
            Next
            Return False
        End Function

        Friend Function MdiChildIsLoaded(ByRef mdiForm As Form, ByVal formName As String, ByVal formText As String) As Boolean
            For Each ChildForm As Form In mdiForm.MdiChildren
                If ChildForm.Name = formName AndAlso ChildForm.Text = formText Then
                    Return True
                End If
            Next
            Return False
        End Function

        Friend Function MdiChildGetInstance(ByRef mdiForm As Form, ByVal formName As String) As Form
            For Each ChildForm As Form In mdiForm.MdiChildren
                If ChildForm.Name = formName Then
                    Return ChildForm
                End If
            Next
            Return Nothing
        End Function

        Friend Function MdiChildGetInstance(ByRef mdiForm As Form, ByVal formName As String, ByVal formText As String) As Form
            For Each ChildForm As Form In mdiForm.MdiChildren
                If ChildForm.Name = formName AndAlso ChildForm.Text = formText Then
                    Return ChildForm
                End If
            Next
            Return Nothing
        End Function

        Friend Sub MdiChildCloseAll(ByRef mdiForm As Form, ParamArray exceptForms() As String)
            For Each FormCurrent As Form In mdiForm.MdiChildren()
                If Not exceptForms.Contains(FormCurrent.Name) Then
                    FormCurrent.Close()
                    FormCurrent.Dispose()
                End If
            Next
        End Sub

        Friend Function IsLoaded(ByVal FormName As String) As Boolean
            For Each CurrentForm As Form In Application.OpenForms
                If CurrentForm.Name = FormName Then
                    Return True
                End If
            Next
            Return False
        End Function

        Friend Function GetInstance(ByVal FormName As String) As Form
            For Each CurrentForm As Form In Application.OpenForms
                If CurrentForm.Name = FormName Then
                    Return CurrentForm
                End If
            Next
            Return Nothing
        End Function

        Friend Sub CloseAll(ParamArray ExceptForms() As String)
            For Each FormCurrent As Form In Application.OpenForms
                If Not ExceptForms.Contains(FormCurrent.Name) Then
                    FormCurrent.Close()
                End If
            Next
        End Sub

        Friend Sub CenterToParent(ByRef ParentForm As Form, ByRef ChildForm As Form)
            With ParentForm
                ChildForm.Top = .Top + CInt((.Height - ChildForm.Height) / 2)
                ChildForm.Left = .Left + CInt((.Width - ChildForm.Width) / 2)
            End With
        End Sub

        Friend Sub ControlsChangeStateEnabled(ByRef ControlsContainer As System.Windows.Forms.Control.ControlCollection, ByVal ValueState As Boolean, ByVal ApplyToLabels As Boolean, ByVal ApplyToPanels As Boolean, ByVal Recursive As Boolean, ByVal ParamArray ControlsExcept() As Object)
            Dim ControlCurrent As Control
            Dim ControlName As Object
            Dim ExceptCurrent As Boolean

            For Each ControlCurrent In ControlsContainer
                If Not ControlsExcept.Contains(ControlCurrent.Name) Then
                    If Recursive AndAlso ControlCurrent.HasChildren Then
                        ControlsChangeStateEnabled(ControlCurrent.Controls, ValueState, ApplyToLabels, ApplyToPanels, Recursive, ControlsExcept)
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.Label Then
                        If ApplyToLabels Then
                            On Error Resume Next
                            ControlCurrent.Enabled = ValueState
                        End If
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.Panel Then
                        If ApplyToPanels Then
                            On Error Resume Next
                            ControlCurrent.Enabled = ValueState
                        End If
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.Button Then
                        If ApplyToPanels Then
                            On Error Resume Next
                            ControlCurrent.Enabled = ValueState
                        End If
                    Else
                        On Error Resume Next
                        ControlCurrent.Enabled = ValueState
                    End If
                End If
            Next ControlCurrent
        End Sub

        Friend Sub ControlsChangeStateReadOnly(ByRef ControlsContainer As System.Windows.Forms.Control.ControlCollection, ByVal ValueState As Boolean, ByVal Recursive As Boolean, ByVal ParamArray ControlsExcept() As String)
            Dim ControlCurrent As Control

            For Each ControlCurrent In ControlsContainer
                If Not ControlsExcept.Contains(ControlCurrent.Name) Then
                    If Recursive AndAlso ControlCurrent.HasChildren Then
                        ControlsChangeStateReadOnly(ControlCurrent.Controls, ValueState, Recursive, ControlsExcept)
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.TextBox Then
                        CType(ControlCurrent, System.Windows.Forms.TextBox).ReadOnly = ValueState
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.MaskedTextBox Then
                        CType(ControlCurrent, System.Windows.Forms.MaskedTextBox).ReadOnly = ValueState
                    ElseIf TypeOf (ControlCurrent) Is Windows.Forms.ComboBox Then
                        ControlCurrent.Enabled = Not ValueState
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.CheckBox Then
                        ControlCurrent.Enabled = Not ValueState
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.DateTimePicker Then
                        ControlCurrent.Enabled = Not ValueState
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.Button Then
                        ControlCurrent.Enabled = Not ValueState
                    ElseIf TypeOf (ControlCurrent) Is System.Windows.Forms.ToolStrip Then
                        ControlCurrent.Enabled = Not ValueState
                    End If
                End If
            Next ControlCurrent
        End Sub
    End Module
End Namespace