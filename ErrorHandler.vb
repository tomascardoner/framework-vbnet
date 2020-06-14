Namespace CardonerSistemas

    Module ErrorHandler
        Friend Const ActiveXCreateError As Integer = -2146233088

        Friend Sub ProcessError(ByRef Exception As Exception, Optional ByVal FriendlyMessageText As String = "", Optional ByVal ShowMessageBox As Boolean = True)
            Dim ExceptionMessageText As String
            Dim MessageTextToLog As String
            Dim formErrorMessageBox As CardonerSistemas.ErrorHandlerMessageBox
            Dim InnerException As Exception

            ' Formatting
            Const TitleUnderline As String = "========================="
            Const ColumnSeparator As String = " |#| "
            Const LabelInnerException As String = "Inner exception:"
            Const LabelWhere As String = "Where: "
            Const LabelUserMessage As String = "User message: "
            Const LabelStackTrace As String = "Stack Trace: "
            Const LabelError As String = "Error: "
            Const MessageErrorOccurred As String = "Se ha producido un error."

            Cursor.Current = Cursors.WaitCursor

            ' Prepare Exception Message Text counting for Inner Exceptions
            ExceptionMessageText = Exception.Message
            If Not Exception.InnerException Is Nothing Then
                If Not Exception.InnerException.InnerException Is Nothing Then
                    InnerException = Exception.InnerException.InnerException
                Else
                    InnerException = Exception.InnerException
                End If
                ExceptionMessageText &= $"{vbCrLf}{vbCrLf}{TitleUnderline}{vbCrLf}{LabelInnerException}{vbCrLf}{InnerException.Message}"
            End If

            MessageTextToLog = $"{LabelWhere}{Exception.Source}"

            If FriendlyMessageText.Length > 0 Then
                MessageTextToLog &= $"{ColumnSeparator}{LabelUserMessage}{FriendlyMessageText}"
            End If

            MessageTextToLog &= $"{ColumnSeparator}{LabelStackTrace}{Exception.StackTrace}{ColumnSeparator}{LabelError}{ExceptionMessageText}"

            My.Application.Log.WriteException(Exception, TraceEventType.Error, ExceptionMessageText)

            Cursor.Current = Cursors.Default

            If ShowMessageBox Then
                If pAppearanceConfig.UseCustomDialogForErrorMessage Then
                    formErrorMessageBox = New CardonerSistemas.ErrorHandlerMessageBox
                    With formErrorMessageBox
                        .lblFriendlyMessage.Text = FriendlyMessageText
                        .lblSourceData.Text = Exception.Source
                        .txtStackTraceData.Text = Exception.StackTrace
                        .txtMessageData.Text = ExceptionMessageText
                        .ShowDialog()
                    End With
                Else
                    MsgBox($"{MessageErrorOccurred}{vbCrLf}{vbCrLf}{FriendlyMessageText}{vbCrLf}{vbCrLf}{ExceptionMessageText}", MsgBoxStyle.Critical, My.Application.Info.Title)
                End If
            End If
        End Sub
    End Module

End Namespace