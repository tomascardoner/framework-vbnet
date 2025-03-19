Namespace CardonerSistemas.Controls

    Module RadioButton

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

        Private Sub ResizeContainer(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, lastRadioButtonBottom As Integer)
            If lastRadioButtonBottom > containter.ClientSize.Height Then
                Dim containerDistanceToForm As Integer
                Dim containerNewHeight As Integer

                form.SuspendLayout()
                containter.SuspendLayout()

                containerDistanceToForm = form.Height - (containter.Top + containter.Height)
                If lastRadioButtonBottom > maximumContainerHeight Then
                    containerNewHeight = maximumContainerHeight
                Else
                    containerNewHeight = containter.Margin.Vertical + lastRadioButtonBottom
                End If
                form.Height = containter.Top + containerNewHeight + containerDistanceToForm
                containter.Height = containerNewHeight

                containter.ResumeLayout()
                form.ResumeLayout()
            End If
        End Sub

        ''' <summary>
        ''' Creates an array of radioButtones based on count parameter
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="containter"></param>
        ''' <param name="maximumContainerHeight"></param>
        ''' <param name="count"></param>
        ''' <param name="clickHandler"></param>
        ''' <param name="widthValue">Sets the width of every radioButton</param>
        ''' <param name="heightValue"></param>
        Friend Sub CreateArray(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, count As Integer, clickHandler As EventHandler, widthValue As Integer, heightValue As Integer, autoSize As Boolean, appearance As Windows.Forms.Appearance, Optional reSizeContainerAndForm As Boolean = True)
            Dim radioButton As System.Windows.Forms.RadioButton = Nothing

            If count = 0 Then
                Return
            End If

            form.SuspendLayout()
            containter.SuspendLayout()

            For index As Integer = 0 To count - 1
                radioButton = New Windows.Forms.RadioButton With {
                    .AutoSize = autoSize,
                    .MinimumSize = New Size(widthValue, heightValue),
                    .Appearance = appearance,
                    .TextAlign = ContentAlignment.MiddleCenter
                }
                containter.Controls.Add(radioButton)
            Next

            containter.ResumeLayout()
            form.ResumeLayout()

            If reSizeContainerAndForm Then
                ResizeContainer(form, containter, maximumContainerHeight, radioButton.Bounds.Bottom)
            End If

            AddHandler radioButton.Click, clickHandler
        End Sub

        ''' <summary>
        ''' Creates an array of radioButtones from a list
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="containter"></param>
        ''' <param name="maximumContainerHeight"></param>
        ''' <param name="list"></param>
        ''' <param name="clickHandler"></param>
        ''' <param name="widthSizeMode"></param>
        ''' <param name="widthValue">Sets the width of every radioButton when Fixed size specified or the minimum width when AutoSizeByTextIndividually or AutoSizeByTextToWidest</param>
        ''' <param name="heightValue"></param>
        Friend Sub CreateArray(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, list As List(Of ValueForList), clickHandler As EventHandler, widthSizeMode As WidthSizeModes, widthValue As Integer, heightValue As Integer, appearance As Windows.Forms.Appearance)
            Dim radioButton As System.Windows.Forms.RadioButton = Nothing
            Dim widestSize As Integer

            If list Is Nothing OrElse list.Count = 0 Then
                Return
            End If

            CreateArray(form, containter, maximumContainerHeight, list.Count, clickHandler, widthValue, heightValue, widthSizeMode = WidthSizeModes.AutoSizeBySystem, appearance, False)

            form.SuspendLayout()
            containter.SuspendLayout()

            For index As Integer = 0 To list.Count - 1
                radioButton = CType(containter.Controls(index), System.Windows.Forms.RadioButton)
                radioButton.Tag = list(index).IdValue
                radioButton.Text = list(index).DisplayValue
                radioButton.MinimumSize = New Size(widthValue, heightValue)
                Select Case widthSizeMode
                    Case WidthSizeModes.AutoSizeBySystem
                        radioButton.AutoSize = True
                    Case WidthSizeModes.AutoSizeByTextIndividually, WidthSizeModes.AutoSizeByTextToWidest
                        Dim extends As Integer

                        extends = 2 + radioButton.Margin.Horizontal + CS_String.GetExtends(radioButton.CreateGraphics(), radioButton.Text, radioButton.Font)
                        If extends > widthValue Then
                            radioButton.Width = extends
                        Else
                            radioButton.Width = widthValue
                        End If
                        If extends > widestSize Then
                            widestSize = extends
                        End If
                End Select
            Next index

            If widthSizeMode = WidthSizeModes.AutoSizeByTextToWidest AndAlso widestSize > 0 Then
                For index As Integer = 0 To list.Count - 1
                    CType(containter.Controls(index), System.Windows.Forms.RadioButton).Width = widestSize
                Next index
            End If

            containter.ResumeLayout()
            form.ResumeLayout()

            ResizeContainer(form, containter, maximumContainerHeight, radioButton.Bounds.Bottom)
        End Sub

    End Module

End Namespace