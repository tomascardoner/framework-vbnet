Namespace CardonerSistemas.Office.Excel

    Module EarlyBinding

        Friend Function InitApplication(ByRef app As Microsoft.Office.Interop.Excel.Application) As Boolean
            Try
                app = New Microsoft.Office.Interop.Excel.Application
                Return True

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al iniciar una instancia de Microsoft Excel.")
                Return False
            End Try
        End Function

    End Module

End Namespace