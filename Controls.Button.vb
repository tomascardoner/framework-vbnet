Namespace CardonerSistemas.Controls

    Module Button

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

        Private Sub ResizeContainer(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, lastButtonBottom As Integer)
            If lastButtonBottom > containter.ClientSize.Height Then
                Dim containerDistanceToForm As Integer
                Dim containerNewHeight As Integer

                form.SuspendLayout()
                containter.SuspendLayout()

                containerDistanceToForm = form.Height - (containter.Top + containter.Height)
                If lastButtonBottom > maximumContainerHeight Then
                    containerNewHeight = maximumContainerHeight
                Else
                    containerNewHeight = containter.Margin.Vertical + lastButtonBottom
                End If
                form.Height = containter.Top + containerNewHeight + containerDistanceToForm
                containter.Height = containerNewHeight

                containter.ResumeLayout()
                form.ResumeLayout()
            End If
        End Sub

        ''' <summary>
        ''' Creates an array of buttons based on count parameter
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="containter"></param>
        ''' <param name="maximumContainerHeight"></param>
        ''' <param name="count"></param>
        ''' <param name="clickHandler"></param>
        ''' <param name="widthValue">Sets the width of every button</param>
        ''' <param name="heightValue"></param>
        Friend Sub CreateArray(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, count As Integer, clickHandler As EventHandler, widthValue As Integer, heightValue As Integer, autoSize As Boolean)
            Dim button As System.Windows.Forms.Button

            If count = 0 Then
                Return
            End If

            form.SuspendLayout()
            containter.SuspendLayout()

            For index As Integer = 0 To count - 1
                button = New Windows.Forms.Button With {
                    .AutoSize = autoSize,
                    .MinimumSize = New Size(widthValue, heightValue)
                }
                containter.Controls.Add(button)
                AddHandler button.Click, clickHandler
            Next

            containter.ResumeLayout()
            form.ResumeLayout()

            ResizeContainer(form, containter, maximumContainerHeight, button.Bounds.Bottom)
            
            AddHandler button.Click, clickHandler
        End Sub

        ''' <summary>
        ''' Creates an array of buttons from a list
        ''' </summary>
        ''' <param name="form"></param>
        ''' <param name="containter"></param>
        ''' <param name="maximumContainerHeight"></param>
        ''' <param name="list"></param>
        ''' <param name="clickHandler"></param>
        ''' <param name="widthSizeMode"></param>
        ''' <param name="widthValue">Sets the width of every button when Fixed size specified or the minimum width when AutoSizeByTextIndividually or AutoSizeByTextToWidest</param>
        ''' <param name="heightValue"></param>
        Friend Sub CreateArray(ByRef form As Form, ByRef containter As System.Windows.Forms.Panel, maximumContainerHeight As Integer, list As List(Of ValueForList), clickHandler As EventHandler, widthSizeMode As WidthSizeModes, widthValue As Integer, heightValue As Integer)
            Dim button As System.Windows.Forms.Button
            Dim widestSize As Integer

            If list Is Nothing OrElse list.Count = 0 Then
                Return
            End If

            CreateArray(form, containter, maximumContainerHeight, list.Count, clickHandler, widthValue, heightValue, widthSizeMode = WidthSizeModes.AutoSizeBySystem)

            form.SuspendLayout()
            containter.SuspendLayout()

            For index As Integer = 0 To list.Count - 1
                button = CType(containter.Controls(index), System.Windows.Forms.Button)
                button.Tag = list(index).IdValue
                button.Text = list(index).DisplayValue
                button.MinimumSize = New Size(widthValue, heightValue)
                Select Case widthSizeMode
                    Case WidthSizeModes.AutoSizeBySystem
                        button.AutoSizeMode = AutoSizeMode.GrowAndShrink
                        button.AutoSize = True
                    Case WidthSizeModes.AutoSizeByTextIndividually, WidthSizeModes.AutoSizeByTextToWidest
                        Dim extends As Integer

                        extends = 2 + button.Margin.Horizontal + CS_String.GetExtends(button.CreateGraphics(), button.Text, button.Font)
                        If extends > widthValue Then
                            button.Width = extends
                        Else
                            button.Width = widthValue
                        End If
                        If extends > widestSize Then
                            widestSize = extends
                        End If
                End Select
            Next index

            If widthSizeMode = WidthSizeModes.AutoSizeByTextToWidest And widestSize > 0 Then
                For index As Integer = 0 To list.Count - 1
                    CType(containter.Controls(index), System.Windows.Forms.Button).Width = widestSize
                Next index
            End If

            containter.ResumeLayout()
            form.ResumeLayout()

            ResizeContainer(form, containter, maximumContainerHeight, button.Bounds.Bottom)
        End Sub

    End Module

End Namespace