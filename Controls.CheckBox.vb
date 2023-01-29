Namespace CardonerSistemas.Controls

    Module CheckBox

        Friend Enum WidthSizeModes
            Fixed
            AutoSizeBySystem
            AutoSizeByTextIndividually
            AutoSizeByTextToWidest
        End Enum

        Friend Class ValueForList
            Friend IdValue As String
            Friend DisplayValue As String
        End Class

        Private Sub ResizeContainer(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, lastCheckBoxBottom As Integer)
            If lastCheckBoxBottom > containter.ClientSize.Height Then
                Dim containerDistanceToForm As Integer
                Dim containerNewHeight As Integer

                form.SuspendLayout()
                containter.SuspendLayout()

                containerDistanceToForm = form.Height - (containter.Top + containter.Height)
                If lastCheckBoxBottom > maximumContainerHeight Then
                    containerNewHeight = maximumContainerHeight
                Else
                    containerNewHeight = containter.Margin.Vertical + lastCheckBoxBottom
                End If
                form.Height = containter.Top + containerNewHeight + containerDistanceToForm
                containter.Height = containerNewHeight

                containter.ResumeLayout()
                form.ResumeLayout()
            End If
        End Sub

        ''' <summary>
        ''' Creates an array of checkboxes based on count parameter
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="containter"></param>
        ''' <param name="maximumContainerHeight"></param>
        ''' <param name="count"></param>
        ''' <param name="clickHandler"></param>
        ''' <param name="widthValue">Sets the width of every checkbox</param>
        ''' <param name="heightValue"></param>
        Friend Sub CreateArray(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, count As Integer, clickHandler As EventHandler, widthValue As Integer, heightValue As Integer, autoSize As Boolean, appearance As Windows.Forms.Appearance)
            Dim checkBox As System.Windows.Forms.CheckBox

            If count = 0 Then
                Return
            End If

            form.SuspendLayout()
            containter.SuspendLayout()

            For index As Integer = 0 To count - 1
                checkBox = New Windows.Forms.CheckBox With {
                    .AutoSize = autoSize,
                    .MinimumSize = New Size(widthValue, heightValue),
                    .Appearance = appearance
                }
                containter.Controls.Add(checkBox)
            Next

            containter.ResumeLayout()
            form.ResumeLayout()

            ResizeContainer(form, containter, maximumContainerHeight, checkbox.Bounds.Bottom)
            
            AddHandler checkBox.Click, clickHandler
        End Sub

        ''' <summary>
        ''' Creates an array of checkboxes from a list
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="containter"></param>
        ''' <param name="maximumContainerHeight"></param>
        ''' <param name="list"></param>
        ''' <param name="clickHandler"></param>
        ''' <param name="widthSizeMode"></param>
        ''' <param name="widthValue">Sets the width of every checkbox when Fixed size specified or the minimum width when AutoSizeByTextIndividually or AutoSizeByTextToWidest</param>
        ''' <param name="heightValue"></param>
        Friend Sub CreateArray(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, list As List(Of ValueForList), clickHandler As EventHandler, widthSizeMode As WidthSizeModes, widthValue As Integer, heightValue As Integer, appearance As Windows.Forms.Appearance)
            Dim checkBox As System.Windows.Forms.CheckBox
            Dim widestSize As Integer

            If list Is Nothing OrElse list.Count = 0 Then
                Return
            End If

            CreateArray(form, containter, maximumContainerHeight, list.Count, clickHandler, widthValue, heightValue, widthSizeMode = WidthSizeModes.AutoSizeBySystem, appearance)

            form.SuspendLayout()
            containter.SuspendLayout()

            For index As Integer = 0 To list.Count - 1
                checkBox = CType(containter.Controls(index), System.Windows.Forms.CheckBox)
                checkBox.Tag = list(index).IdValue
                checkBox.Text = list(index).DisplayValue
                checkBox.MinimumSize = New Size(widthValue, heightValue)
                Select Case widthSizeMode
                    Case WidthSizeModes.AutoSizeBySystem
                        checkBox.AutoSize = True
                    Case WidthSizeModes.AutoSizeByTextIndividually, WidthSizeModes.AutoSizeByTextToWidest
                        Dim extends As Integer

                        extends = 2 + checkBox.Margin.Horizontal + CS_String.GetExtends(checkBox.CreateGraphics(), checkBox.Text, checkBox.Font)
                        If extends > widthValue Then
                            checkBox.Width = extends
                        Else
                            checkBox.Width = widthValue
                        End If
                        If extends > widestSize Then
                            widestSize = extends
                        End If
                End Select
            Next index

            If widthSizeMode = WidthSizeModes.AutoSizeByTextToWidest And widestSize > 0 Then
                For index As Integer = 0 To list.Count - 1
                    CType(containter.Controls(index), System.Windows.Forms.CheckBox).Width = widestSize
                Next index
            End If

            containter.ResumeLayout()
            form.ResumeLayout()

            ResizeContainer(form, containter, maximumContainerHeight, checkBox.Bounds.Bottom)
        End Sub

    End Module

End Namespace