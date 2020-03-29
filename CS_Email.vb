Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Net.Mail

Module CS_Email
    Private Const EmailValidationRegularExpression As String = "^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z_])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$"

    Private invalid As Boolean = False

    Public Function IsValidEmail(ByVal Value As String, Optional ByVal RegularExpression As String = EMAIL_VALIDATION_REGULAREXPRESSION) As Boolean
        invalid = False
        If String.IsNullOrEmpty(Value) Then Return False

        ' Use IdnMapping class to convert Unicode domain names. 
        Try
            Value = Regex.Replace(Value, "(@)(.+)$", AddressOf DomainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(200))
        Catch e As RegexMatchTimeoutException
            Return False
        End Try

        If invalid Then Return False

        ' Return true if strIn is in valid e-mail format. 
        Try
            Return Regex.IsMatch(Value, RegularExpression, RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(250))

        Catch e As RegexMatchTimeoutException
            Return False
        End Try
    End Function

    Private Function DomainMapper(match As Match) As String
        ' IdnMapping class with default property values. 
        Dim idn As New IdnMapping()

        Dim domainName As String = match.Groups(2).Value
        Try
            domainName = idn.GetAscii(domainName)
        Catch e As ArgumentException
            invalid = True
        End Try
        Return match.Groups(1).Value + domainName
    End Function

    Friend Function Send() As Boolean
        'Dim mail As New MailMessage()

        'mail.From = New MailAddress("info@cholilasrl.com.ar")
        'mail.[To].("tomas@cardoner.com.ar")

        ''set the content
        'mail.Subject = String.Format("Semanal {0}-{1}", Anio, Mes)
        'mail.Body = ""

        ''set the server
        'Dim smtp As New SmtpClient("mail.cholilasrl.com.ar")

        'smtp.Credentials = New System.Net.NetworkCredential("info@cholilasrl.com.ar", "")

        ''Attachments
        'mail.Attachments.Add(New System.Net.Mail.Attachment(AttachmentFilename))

        ''send the message
        'Try
        '    smtp.Send(mail)
        '    smtp.Timeout
        '    'Response.Write("Your Email has been sent sucessfully - Thank You")
        'Catch exc As Exception
        '    'Response.Write("Send failure: " & exc.ToString())
        'End Try
        Return True
    End Function
End Module
