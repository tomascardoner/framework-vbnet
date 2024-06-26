Imports System.IO

Module CS_FileLog
    Friend Enum LogEntryType
        Information
        Warning
        Critical
    End Enum

    Private Const COLUMN_DELIMITER As String = " |#| "

    Friend Function GetLogEntryTypeName(ByVal EntryType As LogEntryType) As String
        Select Case EntryType
            Case LogEntryType.Information
                Return "INF"
            Case LogEntryType.Warning
                Return "WRN"
            Case LogEntryType.Critical
                Return "ERR"
            Case Else
                Return String.Empty
        End Select
    End Function

    Private Sub Write_Internal(ByVal LogFolder As String, ByVal LogFileName As String, ByVal EntryType As LogEntryType, ByVal LogMessage As String, ByVal WriteLine As Boolean)
        If LogFolder <> String.Empty AndAlso LogFileName <> String.Empty Then
            Try
                If Not LogFolder.EndsWith("\") Then
                    LogFolder &= "\"
                End If
                If Not Directory.Exists(LogFolder) Then
                    Directory.CreateDirectory(LogFolder)
                End If

                Using sw As New StreamWriter(File.Open(LogFolder & LogFileName, FileMode.Append))
                    If WriteLine Then
                        sw.WriteLine(DateTime.Now.ToString & COLUMN_DELIMITER & GetLogEntryTypeName(EntryType) & COLUMN_DELIMITER & LogMessage)
                    Else
                        sw.Write(DateTime.Now.ToString & COLUMN_DELIMITER & GetLogEntryTypeName(EntryType) & COLUMN_DELIMITER & LogMessage)
                    End If
                End Using

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al escribir el mensaje de registro (Log).")
            End Try
        End If
    End Sub

    Friend Sub Write(logFolder As String, logFileName As String, entryType As LogEntryType, logMessage As String)
        Write_Internal(logFolder, logFileName, entryType, logMessage, False)
    End Sub

    Friend Sub WriteLine(logFolder As String, logFileName As String, entryType As LogEntryType, logMessage As String)
        Write_Internal(logFolder, logFileName, entryType, logMessage, True)
    End Sub

    Friend Sub Write(logEnabled As Boolean, logFolder As String, logFileName As String, entryType As LogEntryType, logMessage As String)
        If logEnabled Then
            Write_Internal(logFolder, logFileName, entryType, logMessage, False)
        End If
    End Sub

    Friend Sub WriteLine(logEnabled As Boolean, logFolder As String, logFileName As String, entryType As LogEntryType, logMessage As String)
        If logEnabled Then
            Write_Internal(logFolder, logFileName, entryType, logMessage, True)
        End If
    End Sub
End Module
