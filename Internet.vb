Imports System.IO
Imports System.Net

Namespace CardonerSistemas
    Module Internet

        Friend Function GetImageFromUrl(ByVal url As String, ByRef image As Image) As Boolean
            Dim webClient As WebClient = New WebClient

            Try
                image = Bitmap.FromStream(New MemoryStream(webClient.DownloadData(url)))
                Return True

            Catch ex As Exception
                CardonerSistemas.ErrorHandler.ProcessError(ex, "Error al obtener la imagen desde internet.")
                Return False
            End Try
        End Function

    End Module
End Namespace