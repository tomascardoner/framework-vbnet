Imports System.IO
Imports System.Text.Json

Namespace CardonerSistemas

    Module ConfigurationJson

        Private Function CheckFileExist(ByVal configFolder As String, ByVal fileName As String) As Boolean
            If File.Exists(Path.Combine(configFolder, fileName)) Then
                Return True
            Else
                MsgBox(String.Format("No se encontró el archivo de configuración '{1}', el cual debe estar ubicado dentro de la carpeta '{0}'.", configFolder, fileName), MsgBoxStyle.Critical, My.Application.Info.Title)
                Return False
            End If
        End Function

        Friend Function LoadFile(Of T)(ByVal configFolder As String, ByVal fileName As String, ByRef config As T) As Boolean
            If Not CheckFileExist(configFolder, fileName) Then
                Return False
            End If

            Dim jsonConfigFileString As String

            Try
                jsonConfigFileString = File.ReadAllText(Path.Combine(configFolder, fileName))
            Catch ex As Exception
                CardonerSistemas.ProcessError(ex, String.Format("Error al leer el archivo de configuración '{0}'.", fileName))
                Return False
            End Try

            Try
                config = JsonSerializer.Deserialize(Of T)(jsonConfigFileString)
            Catch ex As Exception
                CardonerSistemas.ProcessError(ex, String.Format("Error al interpretar el archivo de configuración '{0}'.", fileName))
                Return False
            End Try

            Return True

        End Function

    End Module

End Namespace
