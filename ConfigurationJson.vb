Imports System.IO
Imports System.Text.Json

Namespace CardonerSistemas

    Module ConfigurationJson

        Private Function CheckFileExist(ByVal configFolder As String, ByVal fileName As String) As Boolean
            If File.Exists(Path.Combine(configFolder, fileName)) Then
                Return True
            Else
                MessageBox.Show($"No se encontró el archivo de configuración '{fileName}', el cual debe estar ubicado dentro de la carpeta '{configFolder}'.", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                CardonerSistemas.ProcessError(ex, $"Error al leer el archivo de configuración '{fileName}'.")
                Return False
            End Try

            Try
                config = JsonSerializer.Deserialize(Of T)(jsonConfigFileString)
            Catch ex As Exception
                CardonerSistemas.ProcessError(ex, $"Error al interpretar el archivo de configuración '{fileName}'.")
                Return False
            End Try

            Return True

        End Function

        Friend Function SaveFile(Of T)(configFolder As String, ByVal fileName As String, ByRef configObject As T, Optional writeIndented As Boolean = True) As Boolean

            Dim jsonConfigFileString As String

            Try
                jsonConfigFileString = JsonSerializer.Serialize(Of T)(configObject, New JsonSerializerOptions() With {.WriteIndented = writeIndented})
            Catch ex As System.Exception
                Dim message As String = $"Error al serializar el objeto en archivo de configuración.{vbCrLf}{vbCrLf}{ex.Message}"
                If ex.InnerException IsNot Nothing Then
                    message += $"{vbCrLf}{vbCrLf}Inner message:{vbCrLf}{ex.InnerException.Message}"
                End If
                MessageBox.Show(message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try

            Try
                File.WriteAllText(Path.Combine(configFolder, fileName), jsonConfigFileString)
            Catch ex As System.Exception
                Dim message As String = $"Error al guardar el archivo de configuración {fileName}.{vbCrLf}{vbCrLf}{ex.Message}"
                If ex.InnerException IsNot Nothing Then
                    message += $"{vbCrLf}{vbCrLf}Inner message:{vbCrLf}{ex.InnerException.Message}"
                End If
                MessageBox.Show(message, My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try

            Return True
        End Function

    End Module

End Namespace