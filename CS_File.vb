Module CS_File
    Friend Function ProcessFilename(ByVal Value As String) As String
        Dim ProcessedString As String

        Const DELIMITER_CHAR As Char = "%"c

        ProcessedString = Value
        ProcessedString = ProcessedString.Replace(DELIMITER_CHAR & "DateTimeUniversalNoSlashes" & DELIMITER_CHAR, DateTime.Now.ToString("yyyyMMdd_HHmmss"))

        Return ProcessedString
    End Function

    Friend Function RemoveInvalidFileNameChars(ByVal UserInput As String) As String
        For Each invalidChar In IO.Path.GetInvalidFileNameChars
            UserInput = UserInput.Replace(invalidChar, "")
        Next
        Return UserInput
    End Function
End Module
