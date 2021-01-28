Imports System.Text.Json
Imports Microsoft.Win32

Namespace CardonerSistemas

    Friend Module Files

        Friend Function GetFileNameFromFullPath(ByVal fullPath As String) As String
            Dim LastSeparator As Integer

            LastSeparator = fullPath.LastIndexOf("\")

            If LastSeparator < 0 Then
                Return fullPath
            End If
            If LastSeparator >= fullPath.Length - 1 Then
                Return String.Empty
            End If

            Return fullPath.Substring(LastSeparator + 1)
        End Function

        Friend Function ProcessFilename(ByVal Value As String) As String
            Dim ProcessedString As String

            Const DelimiterCharacter As Char = "%"c

            ProcessedString = Value
            ProcessedString = ProcessedString.Replace(DelimiterCharacter & "DateTimeUniversalNoSlashes" & DelimiterCharacter, System.DateTime.Now.ToString("yyyyMMdd_HHmmss"))

            Return ProcessedString
        End Function

        Friend Function RemoveInvalidFileNameChars(ByVal UserInput As String) As String
            For Each invalidChar In IO.Path.GetInvalidFileNameChars
                UserInput = UserInput.Replace(invalidChar, "")
            Next
            Return UserInput
        End Function

#Region "Process folder name"

        Private Const FolderTagDropbox As String = "{DROPBOX}"
        Private Const FolderTagGoogleDrive As String = "{GOOGLEDRIVE}"
        Private Const FolderTagOneDrive As String = "{ONEDRIVE}"
        Private Const FolderTagICloudDrive As String = "{ICLOUDDRIVE}"

        Friend Function ProcessFolderName(ByVal folderName As String) As String
            Dim folderNameProcessed As String = folderName

            If folderNameProcessed Is Nothing Then
                Return ""
            End If

            ' Replace Dropbox path
            If folderNameProcessed.Contains(FolderTagDropbox) Then
                Dim dropboxFolder As String = ""
                If GetDropboxPath(dropboxFolder) Then
                    folderNameProcessed = folderName.Replace(FolderTagDropbox, dropboxFolder).Trim()
                End If
            End If

            ' Replace Google Drive path
            If folderName.Contains(FolderTagGoogleDrive) Then
                Dim googleDriveFolder As String = ""
                If GetGoogleDrivePath(googleDriveFolder) Then
                    folderNameProcessed = folderName.Replace(FolderTagGoogleDrive, googleDriveFolder).Trim()
                End If
            End If

            ' Replace OneDrive path
            If folderName.Contains(FolderTagOneDrive) Then
                Dim oneDriveFolder As String = ""
                If GetOneDrivePath(oneDriveFolder) Then
                    folderNameProcessed = folderName.Replace(FolderTagOneDrive, oneDriveFolder).Trim()
                End If
            End If

            Return folderNameProcessed
        End Function

#End Region

#Region "Cloud storage - Dropbox"

        Private Class DropboxConfigInfoRoot
            Public Property personal As DropboxConfigInfoPersonal
        End Class

        Private Class DropboxConfigInfoPersonal
            Public Property path As String
            Public Property host As Long
            Public Property is_team As Boolean
            Public Property subscription_type As String
        End Class

        Friend Function GetDropboxPath(ByRef path As String) As Boolean
            Static folderDropbox As String

            If folderDropbox IsNot Nothing Then
                path = folderDropbox
                Return True
            End If

            Const folderName As String = "Dropbox"
            Const configFilename As String = "info.json"

            Dim applicationDatafolder As String
            Dim configFilePath As String
            Dim configFileString As String
            Dim configInfo As DropboxConfigInfoRoot

            ' Gets the path to the Dropbox config file
            applicationDatafolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            If applicationDatafolder.Length <> 0 Then
                configFilePath = System.IO.Path.Combine(applicationDatafolder, folderName, configFilename)

                If System.IO.File.Exists(configFilePath) Then
                    Try
                        configFileString = System.IO.File.ReadAllText(configFilePath)
                    Catch ex As Exception
                        MessageBox.Show($"Ha ocurrido un error al leer el archivo de configuración de Dropbox ({configFilePath})\n\nError: {ex.Message}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End Try

                    Try
                        configInfo = JsonSerializer.Deserialize(Of DropboxConfigInfoRoot)(configFileString)
                    Catch ex As Exception
                        MessageBox.Show($"Ha ocurrido un error al interpretar el archivo de configuración de Dropbox ({configFilename})\n\nError: {ex.Message}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Return False
                    End Try

                    If configInfo.personal IsNot Nothing AndAlso configInfo.personal.path IsNot Nothing Then
                        path = configInfo.personal.path
                        folderDropbox = path
                        Return True
                    End If
                End If
            End If

            Return False
        End Function

#End Region

#Region "Cloud storage - Google Drive"

        Friend Function GetGoogleDrivePath(ByRef path As String) As Boolean
            Static folderGoogleDrive As String

            If folderGoogleDrive IsNot Nothing Then
                path = folderGoogleDrive
                Return True
            End If

            Try
                Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Google\\Drive")
                    If key IsNot Nothing then
                        Dim value As Object = key.GetValue("Path")
                        If value Isnot Nothing then
                            path = value.ToString()
                            folderGoogleDrive = path
                            Return True
                        End if
                    End if
                End Using
                Return False
            Catch ex As Exception
                MessageBox.Show($"Ha ocurrido un error al obtener la ubicación de Google Drive desde el Registro de Windows.\n\nError: {ex.Message}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

#Region "Cloud storage - OneDrive"

        Friend Function GetOneDrivePath(ByRef path As String) As Boolean
            Static folderOneDrive As String

            If folderOneDrive IsNot Nothing Then
                path = folderOneDrive
                Return True
            End If

            Try
                Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\\Microsoft\\OneDrive")
                    If key IsNot Nothing Then
                        Dim value As Object = key.GetValue("UserFolder")
                        If value IsNot Nothing Then
                            path = value.ToString()
                            folderOneDrive = path
                            Return True
                        End If
                    End If
                End Using
                Return False
            Catch ex As Exception
                MessageBox.Show($"Ha ocurrido un error al obtener la ubicación de OneDrive desde el Registro de Windows.\n\nError: {ex.Message}", My.Application.Info.Title, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End Try
        End Function

#End Region

    End Module

End Namespace