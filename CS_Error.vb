Module CS_Error
    Friend Const ACTIVEX_CREATE_ERROR As Integer = -2146233088

    Friend Sub ProcessError(ByRef Exception As Exception, Optional ByVal FriendlyMessageText As String = "", Optional ByVal ShowMessageBox As Boolean = True)
        Dim ExceptionMessageText As String
        Dim MessageTextToLog As String
        Dim formErrorMessageBox As CS_ErrorMessageBox
        Dim InnerException As Exception

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        ' Prepare Exception Message Text counting for Inner Exceptions
        ExceptionMessageText = Exception.Message
        If Not Exception.InnerException Is Nothing Then
            If Not Exception.InnerException.InnerException Is Nothing Then
                InnerException = Exception.InnerException.InnerException
            Else
                InnerException = Exception.InnerException
            End If
            ExceptionMessageText = ExceptionMessageText & vbCrLf & vbCrLf & "=========================" & vbCrLf & "INNER EXCEPTION:" & vbCrLf & InnerException.Message
        End If

        MessageTextToLog = "Where: " & Exception.Source

        If FriendlyMessageText <> "" Then
            MessageTextToLog = MessageTextToLog & " |#| User Message: " & FriendlyMessageText
        End If

        MessageTextToLog = MessageTextToLog & " |#| Stack Trace: " & Exception.StackTrace & " |#| Error: " & ExceptionMessageText

        My.Application.Log.WriteException(Exception, TraceEventType.Error, ExceptionMessageText)

        System.Windows.Forms.Cursor.Current = Cursors.Default

        If ShowMessageBox Then
            If My.Settings.UseCustomDialogForErrorMessage Then
                formErrorMessageBox = New CS_ErrorMessageBox
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