Namespace CardonerSistemas

    Friend Class SpecialFolders

        Friend Shared Function ProcessString(ByVal Value As String) As String
            Dim ProcessedString As String

            Const DeilimiterChar As Char = "%"c

            If Value.Contains(DeilimiterChar) Then
                ProcessedString = Value
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "ApplicationFolder" & DeilimiterChar, My.Application.Info.DirectoryPath)
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "AdminTools" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.AdminTools))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "ApplicationData" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CDBurning" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CDBurning))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonAdminTools" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonAdminTools))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonApplicationData" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonDesktopDirectory" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonDocuments" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonMusic" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonMusic))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonOemLinks" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonOemLinks))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonPictures" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonPictures))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonProgramFiles" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonProgramFilesX86" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonPrograms" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonStartMenu" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonStartup" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonTemplates" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonTemplates))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "CommonVideos" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.CommonVideos))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Cookies" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Cookies))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Desktop" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "DesktopDirectory" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Favorites" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Favorites))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Fonts" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Fonts))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "History" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.History))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "InternetCache" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.InternetCache))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "LocalApplicationData" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "LocalizedResources" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.LocalizedResources))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "MyComputer" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.MyComputer))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "MyDocuments" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "MyMusic" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.MyMusic))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "MyPictures" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "MyVideos" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.MyVideos))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "NetworkShortcuts" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Personal" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Personal))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "PrinterShortcuts" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.PrinterShortcuts))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "ProgramFiles" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "ProgramFilesX86" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Programs" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Programs))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Recent" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Recent))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Resources" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Resources))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "SendTo" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.SendTo))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "StartMenu" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.StartMenu))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Startup" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Startup))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "System" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.System))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "SystemX86" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.SystemX86))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Templates" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Templates))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "UserProfile" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
                ProcessedString = CS_String.ReplaceString(ProcessedString, DeilimiterChar & "Windows" & DeilimiterChar, Environment.GetFolderPath(Environment.SpecialFolder.Windows))

                Return ProcessedString
            Else
                Return Value
            End If

        End Function

        Friend Shared Function GetGuidTempFileName(fileExtension As String) As String
            Return System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + fileExtension
        End Function

        Friend Shared Function GetTempFileName() As String
            Return System.IO.Path.GetTempFileName()
        End Function

        Friend Shared Function GetTempFileName(fileExtension As String) As String
            Return GetTempFileName().Replace(".tmp", "." + fileExtension)
        End Function

    End Class

End Namespace