Module CS_Form

    Friend Sub MDIChild_Show(ByRef MDIForm As Form, ByRef ChildForm As Form, ByVal CenterForm As Boolean)
        MDIForm.Cursor = Cursors.WaitCursor

        ChildForm.MdiParent = MDIForm
        If CenterForm Then
            CS_Form.CenterToParent(MDIForm, ChildForm)
        Else
            CS_Form.MDIChild_PositionAndSizeToFit(MDIForm, ChildForm)
        End If
        ChildForm.Show()
        If ChildForm.WindowState = FormWindowState.Minimized Then
            ChildForm.WindowState = FormWindowState.Normal
        End If
        ChildForm.Focus()

        MDIForm.Cursor = Cursors.Default
    End Sub

    Friend Sub MDIChild_PositionAndSizeToFit(ByRef MDIForm As Form, ByRef ChildForm As Form)
        With ChildForm
            .SuspendLayout()

            .MdiParent = MDIForm

            If .WindowState <> FormWindowState.Normal Then
                .WindowState = FormWindowState.Normal
            End If

            .Dock = DockStyle.Fill

            .ResumeLayout(True)
        End With
    End Sub

    Friend Sub MDIChild_CenterToClientArea(ByRef MDIForm As Form, ByRef ChildForm As Form, ByVal ClientSize As Drawing.Size)
        With ChildForm
            If .WindowState <> FormWindowState.Normal Then
                .WindowState = FormWindowState.Normal
            End If

            .Top = CInt((ClientSize.Height - .Height) / 2)
            .Left = CInt((ClientSize.Width - .Width) / 2)
        End With
    End Sub

    Friend Function MDIChild_IsLoaded(ByRef MDIForm As Form, ByVal FormName As String) As Boolean
        For Each ChildForm As Form In MDIForm.MdiChildren
            If ChildForm.Name = FormName Then
                Return True
            End If
        Next
        Return False
    End Function

    Friend Function MDIChild_IsLoaded(ByRef MDIForm As Form, ByVal FormName As String, ByVal FormText As String) As Boolean
        For Each ChildForm As Form In MDIForm.MdiChildren
            If ChildForm.Name = FormName And ChildForm.Text = FormText Then
                Return True
            End If
        Next
        Return False
    End Function

    Friend Function MDIChild_GetInstance(ByRef MDIForm As Form, ByVal FormName As String) As Form
        For Each ChildForm As Form In MDIForm.MdiChildren
            If ChildForm.Name = FormName Then
                Return ChildForm
            End If
        Next
        Return Nothing
    End Function

    Friend Function MDIChild_GetInstance(ByRef MDIForm As Form, ByVal FormName As String, ByVal FormText As String) As Form
        For Each ChildForm As Form In MDIForm.MdiChildren
            If ChildForm.Name = FormName And ChildForm.Text = FormText Then
                Return ChildForm
            End If
        Next
        Return Nothing
    End Function

    Friend Sub MDIChild_CloseAll(ByRef MDIForm As Form, ParamArray ExceptForms() As String)
        For Each FormCurrent As Form In MDIForm.MdiChildren
            If Not ExceptForms.Contains(FormCurrent.Name) Then
                FormCurrent.Close()
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
                ElseIf TypeOf (ControlCurrent) Is Label Then
                    If ApplyToLabels Then
                        On Error Resume Next
                        ControlCurrent.Enabled = ValueState
                    End If
                ElseIf TypeOf (ControlCurrent) Is Panel Then
                    If ApplyToPanels Then
                        On Error Resume Next
                        ControlCurrent.Enabled = ValueState
                    End If
                ElseIf TypeOf (ControlCurrent) Is Button Then
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
                ElseIf TypeOf (ControlCurrent) Is TextBox Then
                    CType(ControlCurrent, TextBox).ReadOnly = ValueState
                ElseIf TypeOf (ControlCurrent) Is MaskedTextBox Then
                    CType(ControlCurrent, MaskedTextBox).ReadOnly = ValueState
                ElseIf TypeOf (ControlCurrent) Is ComboBox Then
                    ControlCurrent.Enabled = Not ValueState
                ElseIf TypeOf (ControlCurrent) Is CheckBox Then
                    ControlCurrent.Enabled = Not ValueState
                ElseIf TypeOf (ControlCurrent) Is DateTimePicker Then
                    ControlCurrent.Enabled = Not ValueState
                ElseIf TypeOf (ControlCurrent) Is Button Then
                    ControlCurrent.Enabled = Not ValueState
                ElseIf TypeOf (ControlCurrent) Is ToolStrip Then
                    ControlCurrent.Enabled = Not ValueState
                End If
            End If
        Next ControlCurrent
    End Sub

End Module