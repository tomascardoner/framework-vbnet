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
            ProcessedString = ProcessedString.Replace(DelimiterCharacter & "DateTimeUniversalNoSlashes" & DelimiterCharacter, DateTime.Now.ToString("yyyyMMdd_HHmmss"))

            Return ProcessedString
        End Function

        Friend Function RemoveInvalidFileNameChars(ByVal UserInput As String) As String
            For Each invalidChar In IO.Path.GetInvalidFileNameChars
                UserInput = UserInput.Replace(invalidChar, "")
            Next
            Return UserInput
        End Function

    End Module

End Namespace