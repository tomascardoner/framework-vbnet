Namespace CardonerSistemas

    Module ErrorHandler
        Friend Const ActiveXCreateError As Integer = -2146233088

        Friend Sub ProcessError(ByRef Exception As Exception, Optional ByVal FriendlyMessageText As String = "", Optional ByVal ShowMessageBox As Boolean = True)
            Dim ExceptionMessageText As String
            Dim MessageTextToLog As String
            Dim formErrorMessageBox As CardonerSistemas.ErrorHandlerMessageBox
            Dim InnerException As Exception

            Cursor.Current = Cursors.WaitCursor

            ' Prepare Exception Message Text counting for Inner Exceptions
            ExceptionMessageText = Exception.Message
            If Not Exception.InnerException Is Nothing Then
                If Not Exception.InnerException.InnerException Is Nothing Then
                    InnerException = Exception.InnerException.InnerException
                Else
                    InnerException = Exception.InnerException
                End If
                ExceptionMessageText &= String.Format("{0}{0}{1}{0}INNER EXCEPTION:{0}{2}", vbCrLf, New String("="c, 25), InnerException.Message)
            End If

            MessageTextToLog = "Where: " & Exception.Source

            If FriendlyMessageText <> "" Then
                MessageTextToLog &= String.Format(" |#| User Message: {0}", FriendlyMessageText)
            End If

            MessageTextToLog &= String.Format(" |#| Stack Trace: {0} |#| Error: {1}", Exception.StackTrace, ExceptionMessageText)

            My.Application.Log.WriteException(Exception, TraceEventType.Error, ExceptionMessageText)

            Cursor.Current = Cursors.Default

            If ShowMessageBox Then
                If My.Settings.UseCustomDialogForErrorMessage Then
                    formErrorMessageBox = New CardonerSistemas.ErrorHandlerMessageBox
                    With formErrorMessageBox
                        .lblFriendlyMessage.Text = FriendlyMessageText
                        .lblSourceData.Text = Exception.Source
                        .txtStackTraceData.Text = Exception.StackTrace
                        .txtMessageData.Text = ExceptionMessageText
                        .ShowDialog()
                    End With
                Else
                    MsgBox("Se ha producido un error." & vbCr & vbCr & FriendlyMessageText & vbCrLf & vbCrLf & ExceptionMessageText, MsgBoxStyle.Critical, My.Application.Info.Title)
                End If
            End If
        End Sub
    End Module

End Namespace