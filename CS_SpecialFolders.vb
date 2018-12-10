Module CS_SpecialFolders
    Friend Function ProcessString(ByVal Value As String) As String
        Dim ProcessedString As String

        Const DELIMITER_CHAR As Char = "%"c

        If Value.Contains(DELIMITER_CHAR) Then
            ProcessedString = Value
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "ApplicationFolder" & DELIMITER_CHAR, My.Application.Info.DirectoryPath)
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "AdminTools" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.AdminTools))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "ApplicationData" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CDBurning" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CDBurning))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonAdminTools" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonAdminTools))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonApplicationData" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonDesktopDirectory" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonDesktopDirectory))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonDocuments" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonDocuments))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonMusic" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonMusic))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonOemLinks" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonOemLinks))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonPictures" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonPictures))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonProgramFiles" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFiles))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonProgramFilesX86" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonPrograms" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonPrograms))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonStartMenu" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonStartup" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonStartup))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonTemplates" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonTemplates))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "CommonVideos" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.CommonVideos))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Cookies" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Cookies))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Desktop" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "DesktopDirectory" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Favorites" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Favorites))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Fonts" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Fonts))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "History" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.History))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "InternetCache" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.InternetCache))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "LocalApplicationData" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "LocalizedResources" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.LocalizedResources))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "MyComputer" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.MyComputer))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "MyDocuments" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "MyMusic" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.MyMusic))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "MyPictures" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "MyVideos" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.MyVideos))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "NetworkShortcuts" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.NetworkShortcuts))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Personal" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Personal))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "PrinterShortcuts" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.PrinterShortcuts))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "ProgramFiles" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "ProgramFilesX86" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Programs" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Programs))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Recent" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Recent))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Resources" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Resources))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "SendTo" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.SendTo))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "StartMenu" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.StartMenu))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Startup" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Startup))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "System" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.System))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "SystemX86" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.SystemX86))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Templates" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Templates))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "UserProfile" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.UserProfile))
            ProcessedString = CS_String.ReplaceString(ProcessedString, DELIMITER_CHAR & "Windows" & DELIMITER_CHAR, Environment.GetFolderPath(Environment.SpecialFolder.Windows))

            Return ProcessedString
        Else
            Return Value
        End If

    End Function
End Module
