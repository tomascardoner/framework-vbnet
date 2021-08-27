Imports System.Text.Json
Imports Microsoft.Win32

Namespace CardonerSistemas

    Friend Module Files
        Private Const FieldDelimiterCharacterStart As Char = "{"c
        Private Const FieldDelimiterCharacterEnd As Char = "}"c

#Region "Files"

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

            Const FieldDateTimeUniversalNoSlashes As String = FieldDelimiterCharacterStart & "DateTimeUniversalNoSlashes" & FieldDelimiterCharacterEnd

            ProcessedString = Value
            ProcessedString = ProcessedString.Replace(FieldDateTimeUniversalNoSlashes, System.DateTime.Now.ToString("yyyyMMdd_HHmmss"))

            Return ProcessedString
        End Function

        Friend Function RemoveInvalidFileNameChars(ByVal UserInput As String) As String
            For Each invalidChar In IO.Path.GetInvalidFileNameChars
                UserInput = UserInput.Replace(invalidChar, "")
            Next
            Return UserInput
        End Function

        Friend Function GetGuidTempFileName(fileExtension As String) As String
            Return System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + fileExtension
        End Function

        Friend Function GetTempFileName() As String
            Return System.IO.Path.GetTempFileName()
        End Function

        Friend Function GetTempFileName(fileExtension As String) As String
            Return GetTempFileName().Replace(".tmp", "." + fileExtension)
        End Function

#End Region

#Region "Process folder name"

        Friend Function ProcessFolderName(ByVal folderName As String) As String
            ' System folders
            Const FolderTagApplicationFolder As String = FieldDelimiterCharacterStart & "ApplicationFolder" & FieldDelimiterCharacterEnd
            Const FolderTagAdminTools As String = FieldDelimiterCharacterStart & "AdminTools" & FieldDelimiterCharacterEnd
            Const FolderTagApplicationData As String = FieldDelimiterCharacterStart & "ApplicationData" & FieldDelimiterCharacterEnd
            Const FolderTagCDBurning As String = FieldDelimiterCharacterStart & "CDBurning" & FieldDelimiterCharacterEnd
            Const FolderTagCommonAdminTools As String = FieldDelimiterCharacterStart & "CommonAdminTools" & FieldDelimiterCharacterEnd
            Const FolderTagCommonApplicationData As String = FieldDelimiterCharacterStart & "CommonApplicationData" & FieldDelimiterCharacterEnd
            Const FolderTagCommonDesktopDirectory As String = FieldDelimiterCharacterStart & "CommonDesktopDirectory" & FieldDelimiterCharacterEnd
            Const FolderTagCommonDocuments As String = FieldDelimiterCharacterStart & "CommonDocuments" & FieldDelimiterCharacterEnd
            Const FolderTagCommonMusic As String = FieldDelimiterCharacterStart & "CommonMusic" & FieldDelimiterCharacterEnd
            Const FolderTagCommonOemLinks As String = FieldDelimiterCharacterStart & "CommonOemLinks" & FieldDelimiterCharacterEnd
            Const FolderTagCommonPictures As String = FieldDelimiterCharacterStart & "CommonPictures" & FieldDelimiterCharacterEnd
            Const FolderTagCommonProgramFiles As String = FieldDelimiterCharacterStart & "CommonProgramFiles" & FieldDelimiterCharacterEnd
            Const FolderTagCommonProgramFilesX86 As String = FieldDelimiterCharacterStart & "CommonProgramFilesX86" & FieldDelimiterCharacterEnd
            Const FolderTagCommonPrograms As String = FieldDelimiterCharacterStart & "CommonPrograms" & FieldDelimiterCharacterEnd
            Const FolderTagCommonStartMenu As String = FieldDelimiterCharacterStart & "CommonStartMenu" & FieldDelimiterCharacterEnd
            Const FolderTagCommonStartup As String = FieldDelimiterCharacterStart & "CommonStartup" & FieldDelimiterCharacterEnd
            Const FolderTagCommonTemplates As String = FieldDelimiterCharacterStart & "CommonTemplates" & FieldDelimiterCharacterEnd
            Const FolderTagCommonVideos As String = FieldDelimiterCharacterStart & "CommonVideos" & FieldDelimiterCharacterEnd
            Const FolderTagCookies As String = FieldDelimiterCharacterStart & "Cookies" & FieldDelimiterCharacterEnd
            Const FolderTagDesktop As String = FieldDelimiterCharacterStart & "Desktop" & FieldDelimiterCharacterEnd
            Const FolderTagDesktopDirectory As String = FieldDelimiterCharacterStart & "DesktopDirectory" & FieldDelimiterCharacterEnd
            Const FolderTagFavorites As String = FieldDelimiterCharacterStart & "Favorites" & FieldDelimiterCharacterEnd
            Const FolderTagFonts As String = FieldDelimiterCharacterStart & "Fonts" & FieldDelimiterCharacterEnd
            Const FolderTagHistory As String = FieldDelimiterCharacterStart & "History" & FieldDelimiterCharacterEnd
            Const FolderTagInternetCache As String = FieldDelimiterCharacterStart & "InternetCache" & FieldDelimiterCharacterEnd
            Const FolderTagLocalApplicationData As String = FieldDelimiterCharacterStart & "LocalApplicationData" & FieldDelimiterCharacterEnd
            Const FolderTagLocalizedResources As String = FieldDelimiterCharacterStart & "LocalizedResources" & FieldDelimiterCharacterEnd
            Const FolderTagMyComputer As String = FieldDelimiterCharacterStart & "MyComputer" & FieldDelimiterCharacterEnd
            Const FolderTagMyDocuments As String = FieldDelimiterCharacterStart & "MyDocuments" & FieldDelimiterCharacterEnd
            Const FolderTagMyMusic As String = FieldDelimiterCharacterStart & "MyMusic" & FieldDelimiterCharacterEnd
            Const FolderTagMyPictures As String = FieldDelimiterCharacterStart & "MyPictures" & FieldDelimiterCharacterEnd
            Const FolderTagMyVideos As String = FieldDelimiterCharacterStart & "MyVideos" & FieldDelimiterCharacterEnd
            Const FolderTagNetworkShortcuts As String = FieldDelimiterCharacterStart & "NetworkShortcuts" & FieldDelimiterCharacterEnd
            Const FolderTagPersonal As String = FieldDelimiterCharacterStart & "Personal" & FieldDelimiterCharacterEnd
            Const FolderTagPrinterShortcuts As String = FieldDelimiterCharacterStart & "PrinterShortcuts" & FieldDelimiterCharacterEnd
            Const FolderTagProgramFiles As String = FieldDelimiterCharacterStart & "ProgramFiles" & FieldDelimiterCharacterEnd
            Const FolderTagProgramFilesX86 As String = FieldDelimiterCharacterStart & "ProgramFilesX86" & FieldDelimiterCharacterEnd
            Const FolderTagPrograms As String = FieldDelimiterCharacterStart & "Programs" & FieldDelimiterCharacterEnd
            Const FolderTagRecent As String = FieldDelimiterCharacterStart & "Recent" & FieldDelimiterCharacterEnd
            Const FolderTagResources As String = FieldDelimiterCharacterStart & "Resources" & FieldDelimiterCharacterEnd
            Const FolderTagSendTo As String = FieldDelimiterCharacterStart & "SendTo" & FieldDelimiterCharacterEnd
            Const FolderTagStartMenu As String = FieldDelimiterCharacterStart & "StartMenu" & FieldDelimiterCharacterEnd
            Const FolderTagStartup As String = FieldDelimiterCharacterStart & "Startup" & FieldDelimiterCharacterEnd
            Const FolderTagSystem As String = FieldDelimiterCharacterStart & "System" & FieldDelimiterCharacterEnd
            Const FolderTagSystemX86 As String = FieldDelimiterCharacterStart & "SystemX86" & FieldDelimiterCharacterEnd
            Const FolderTagTemplates As String = FieldDelimiterCharacterStart & "Templates" & FieldDelimiterCharacterEnd
            Const FolderTagUserProfile As String = FieldDelimiterCharacterStart & "UserProfile" & FieldDelimiterCharacterEnd
            Const FolderTagWindows As String = FieldDelimiterCharacterStart & "Windows" & FieldDelimiterCharacterEnd

            ' Cloud drives
            Const FolderTagDropbox As String = FieldDelimiterCharacterStart & "Dropbox" & FieldDelimiterCharacterEnd
            Const FolderTagGoogleDrive As String = FieldDelimiterCharacterStart & "GoogleDrive" & FieldDelimiterCharacterEnd
            Const FolderTagOneDrive As String = FieldDelimiterCharacterStart & "OneDrive" & FieldDelimiterCharacterEnd
            'Const FolderTagICloudDrive As String = FieldDelimiterCharacterStart & "iCloudDrive" & FieldDelimiterCharacterEnd

            Dim folderNameProcessed As String

            If folderName Is Nothing Then
                Return String.Empty
            End If

            ' No hay campos en el nombre de la carpeta, por ende, no hay nada que procesar.
            If Not (folderName.Contains(FieldDelimiterCharacterStart) AndAlso folderName.Contains(FieldDelimiterCharacterEnd)) Then
                Return folderName
            End If

            folderNameProcessed = folderName.Trim()

            ' Reemplazo todas las carpetas de sistema
            folderNameProcessed = folderNameProcessed.Replace(FolderTagApplicationFolder, My.Application.Info.DirectoryPath)
            folderNameProcessed = folderNameProcessed.Replace(FolderTagAdminTools, Environment.GetFolderPath(Environment.SpecialFolder.AdminTools))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagApplicationData, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCDBurning, Environment.GetFolderPath(Environment.SpecialFolder.CDBurning))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonAdminTools, Environment.GetFolderPath(Environment.SpecialFolder.CommonAdminTools))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonApplicationData, Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonDesktopDirectory, Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonDocuments, Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonMusic, Environment.GetFolderPath(Environment.SpecialFolder.CommonMusic))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonOemLinks, Environment.GetFolderPath(Environment.SpecialFolder.CommonOemLinks))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonPictures, Environment.GetFolderPath(Environment.SpecialFolder.CommonPictures))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonProgramFiles, Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonProgramFilesX86, Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonPrograms, Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonStartMenu, Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonStartup, Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonTemplates, Environment.GetFolderPath(Environment.SpecialFolder.CommonTemplates))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCommonVideos, Environment.GetFolderPath(Environment.SpecialFolder.CommonVideos))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagCookies, Environment.GetFolderPath(Environment.SpecialFolder.Cookies))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagDesktop, Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagDesktopDirectory, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagFavorites, Environment.GetFolderPath(Environment.SpecialFolder.Favorites))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagFonts, Environment.GetFolderPath(Environment.SpecialFolder.Fonts))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagHistory, Environment.GetFolderPath(Environment.SpecialFolder.History))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagInternetCache, Environment.GetFolderPath(Environment.SpecialFolder.InternetCache))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagLocalApplicationData, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagLocalizedResources, Environment.GetFolderPath(Environment.SpecialFolder.LocalizedResources))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagMyComputer, Environment.GetFolderPath(Environment.SpecialFolder.MyComputer))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagMyDocuments, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagMyMusic, Environment.GetFolderPath(Environment.SpecialFolder.MyMusic))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagMyPictures, Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagMyVideos, Environment.GetFolderPath(Environment.SpecialFolder.MyVideos))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagNetworkShortcuts, Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagPersonal, Environment.GetFolderPath(Environment.SpecialFolder.Personal))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagPrinterShortcuts, Environment.GetFolderPath(Environment.SpecialFolder.PrinterShortcuts))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagProgramFiles, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagProgramFilesX86, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagPrograms, Environment.GetFolderPath(Environment.SpecialFolder.Programs))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagRecent, Environment.GetFolderPath(Environment.SpecialFolder.Recent))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagResources, Environment.GetFolderPath(Environment.SpecialFolder.Resources))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagSendTo, Environment.GetFolderPath(Environment.SpecialFolder.SendTo))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagStartMenu, Environment.GetFolderPath(Environment.SpecialFolder.StartMenu))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagStartup, Environment.GetFolderPath(Environment.SpecialFolder.Startup))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagSystem, Environment.GetFolderPath(Environment.SpecialFolder.System))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagSystemX86, Environment.GetFolderPath(Environment.SpecialFolder.SystemX86))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagTemplates, Environment.GetFolderPath(Environment.SpecialFolder.Templates))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagUserProfile, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
            folderNameProcessed = folderNameProcessed.Replace(FolderTagWindows, Environment.GetFolderPath(Environment.SpecialFolder.Windows))

            ' Replace Dropbox path
            If folderNameProcessed.Contains(FolderTagDropbox) Then
                Dim dropboxFolder As String = ""
                If GetDropboxPath(dropboxFolder) Then
                    folderNameProcessed = folderNameProcessed.Replace(FolderTagDropbox, dropboxFolder)
                End If
            End If

            ' Replace Google Drive path
            If folderName.Contains(FolderTagGoogleDrive) Then
                Dim googleDriveFolder As String = ""
                If GetGoogleDrivePath(googleDriveFolder) Then
                    folderNameProcessed = folderNameProcessed.Replace(FolderTagGoogleDrive, googleDriveFolder)
                End If
            End If

            ' Replace OneDrive path
            If folderName.Contains(FolderTagOneDrive) Then
                Dim oneDriveFolder As String = ""
                If GetOneDrivePath(oneDriveFolder) Then
                    folderNameProcessed = folderNameProcessed.Replace(FolderTagOneDrive, oneDriveFolder)
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