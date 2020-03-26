Namespace CardonerSistemas

    Module Process

        Friend Sub Start(ByVal processName As String)
            If processName.Length > 0 Then
                Try
                    ' Cursor = Cursors.AppStarting
                    System.Diagnostics.Process.Start(processName)

                Catch ex As Exception
                    ErrorHandler.ProcessError(ex, String.Format("Error al iniciar el proceso '{0}'.", processName))

                Finally
                    'Cursor = Cursors.Default
                End Try
            End If
        End Sub

    End Module
    
End Namespace