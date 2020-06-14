Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml.Serialization

Namespace CardonerSistemas

    Public Class Serializer

        Public Function Deserialize(Of T As Class)(ByVal input As String) As T
            Dim ser As XmlSerializer = New XmlSerializer(GetType(T))

            Using sr As StringReader = New StringReader(input)
                Return CType(ser.Deserialize(sr), T)
            End Using
        End Function

        Public Function Serialize(Of T)(ByVal ObjectToSerialize As T) As String
            Dim xmlSerializer As XmlSerializer = New XmlSerializer(ObjectToSerialize.[GetType]())

            Using textWriter As StringWriter = New StringWriter()
                xmlSerializer.Serialize(textWriter, ObjectToSerialize)
                Return textWriter.ToString()
            End Using
        End Function

    End Class

    Module Configuration

        Private Const ErrorFileBadFormat As Integer = -2146233079
        Private Const ErrorFileBadFormatInnerElement As Integer = -2146232000
        Private Const ErrorFileBadFormatInnerElementValue As Integer = -2146233033
        Private Const ErrorFileBadFormatPositionPattern As String = "\(\d+, \d+\)"
        Private Const ErrorFileBadFormatPositionLinePattern As String = "\d+"

        Private Function CheckFileExist(ByVal configFolder As String, ByVal fileName As String) As Boolean
            If File.Exists(configFolder & fileName) Then
                Return True
            Else
                MsgBox(String.Format("No se encontró el archivo de configuración '{1}', el cual debe estar ubicado dentro de la carpeta '{0}'.", configFolder, fileName), MsgBoxStyle.Critical, My.Application.Info.Title)
                Return False
            End If
        End Function

        Friend Function LoadFile(Of T)(ByVal configFolder As String, ByVal fileName As String, ByRef configObject As T) As Boolean
            Dim serializer As XmlSerializer
            Dim fileStream As FileStream

            If Not CheckFileExist(configFolder, fileName) Then
                Return False
            End If

            Try
                serializer = New XmlSerializer(GetType(T))
                fileStream = New FileStream(configFolder & fileName, FileMode.Open, FileAccess.Read, FileShare.Read)
                configObject = DirectCast(serializer.Deserialize(fileStream), T)
                Return True

            Catch ex As Exception
                Select Case ex.HResult
                    Case ErrorFileBadFormat
                        ' El formato del archivo es incorrecto.
                        If Not ex.InnerException Is Nothing Then
                            ' Intento obtener mayor información del error
                            Select Case ex.InnerException.HResult
                                Case ErrorFileBadFormatInnerElement
                                    ' Error en el elemento
                                    MsgBox(String.Format("Error en el formato del archivo de configuración {1}.{0}{0}{2}", vbCrLf, fileName, ex.InnerException.Message), MsgBoxStyle.Critical, My.Application.Info.Title)
                                Case ErrorFileBadFormatInnerElementValue
                                    ' El valor especificado en el elemento no es válido

                                    ' Trato de obtener la línea en la que se encuentra el error
                                    Dim textoPosicion As String
                                    Dim textoPosicionLinea As String
                                    Dim posicionLinea As Integer
                                    Dim regexPosicion As New Regex(ErrorFileBadFormatPositionPattern)
                                    Dim regexLinea As New Regex(ErrorFileBadFormatPositionLinePattern)

                                    If regexPosicion.IsMatch(ex.Message) Then
                                        textoPosicion = regexPosicion.Match(ex.Message).Value
                                        textoPosicionLinea = regexLinea.Match(textoPosicion).Value
                                        Integer.TryParse(textoPosicionLinea, posicionLinea)
                                    End If

                                    If posicionLinea > 0 Then
                                        MsgBox(String.Format("Error en el formato del archivo de configuración '{1}', en la línea número {2}.{0}{0}{3}", vbCrLf, fileName, posicionLinea - 1, ex.InnerException.Message), MsgBoxStyle.Critical, My.Application.Info.Title)
                                    Else
                                        MsgBox(String.Format("Error en el formato del archivo de configuración '{1}'.{0}{0}{2}", vbCrLf, fileName, ex.InnerException.Message), MsgBoxStyle.Critical, My.Application.Info.Title)
                                    End If
                                Case Else
                                    CardonerSistemas.ProcessError(ex, String.Format("Error al cargar el archivo de configuración '{0}'.", fileName))
                            End Select
                        End If
                    Case Else
                        CardonerSistemas.ProcessError(ex, String.Format("Error al cargar el archivo de configuración '{0}'.", fileName))
                End Select
                serializer = Nothing
                fileStream = Nothing
                Return False
            End Try

        End Function


        'Private Function SaveFile(ByVal configFolder As String) As Boolean
        '    Dim serializer As Serializer = New Serializer()
        '    Dim oututText As String

        '    Try
        '        oututText = serializer.Serialize(Of Config)(pConfig)
        '        serializer = Nothing
        '        Return True

        '    Catch ex As Exception
        '        CardonerSistemas.ProcessError(ex, String.Format("Error al guardar el archivo de configuración '{0}'.", FileName))
        '        serializer = Nothing
        '        Return False
        '    End Try

        'End Function

        Friend Function ConvertStringToFont(ByVal value As String) As Font
            Dim converter As TypeConverter
            Dim convertedFont As Font

            Try
                converter = TypeDescriptor.GetConverter(GetType(Font))
                convertedFont = CType(converter.ConvertFromString(value), Font)
            Catch ex As Exception
                convertedFont = Nothing
            End Try

            Return convertedFont
        End Function

        Friend Function ConvertFontToString(ByVal value As Font) As String
            Dim converter As TypeConverter
            Dim convertedString As String

            Try
                converter = TypeDescriptor.GetConverter(GetType(Font))
                convertedString = converter.ConvertToString(value)
            Catch ex As Exception
                convertedString = ""
            End Try

            Return convertedString
        End Function

    End Module

End Namespace