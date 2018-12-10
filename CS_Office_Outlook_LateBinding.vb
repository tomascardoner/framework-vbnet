Option Strict Off

Module CS_Office_Outlook_LateBinding

#Region "Declarations"
    Friend Enum OlItemType As Integer
        olMailItem = 0
        olAppointmentItem = 1
        olContactItem = 2
        olTaskItem = 3
        olJournalItem = 4
        olNoteItem = 5
        olPostItem = 6
        olDistributionListItem = 7
        olMobileItemSMS = 11
        olMobileItemMMS = 12
    End Enum

    Friend Enum OlAttachmentType As Integer
        olByValue = 1
        olByReference = 4
        olEmbeddeditem = 5
        olOLE = 6
    End Enum

    Friend Enum OlAccountType As Integer
        olExchange = 0
        olImap = 1
        olPop3 = 2
        olHttp = 3
        olEas = 4
        olOtherAccount = 5
    End Enum

    Friend Enum OlBodyFormat As Integer
        olFormatUnspecified = 0
        olFormatPlain = 1
        olFormatHTML = 2
        olFormatRichText = 3
    End Enum

    Friend Class Attachment
        Friend Property DisplayName As String
        Friend Property FileName As String
        Friend Property PathName As String
        Friend Property Position As Integer
        Friend Property Size As Integer
        Friend Property Type As OlAttachmentType
        Friend Property ContentStream As System.IO.Stream

        Friend Sub New()

        End Sub

        Friend Sub New(ByVal Stream As System.IO.Stream, ByVal Name As String)
            ContentStream = Stream
            DisplayName = Name
        End Sub
    End Class

    Friend Class Account
        Friend Property AccountType As OlAccountType
        Friend Property DisplayName As String
        Friend Property SmtpAddress As String
        Friend Property UserName As String
    End Class

    Friend Class MailItem
        Friend Property Attachments As New List(Of Attachment)
        Friend Property CC As String
        Friend Property BCC As String
        Friend Property Body As String
        Friend Property BodyFormat As OlBodyFormat
        Friend Property RTFBody As Object
        Friend Property SendUsingAccount As Account
        Friend Property Subject As String
        Friend Property [To] As String
        Friend Property Unread As Boolean

        Friend Sub New()
            BodyFormat = OlBodyFormat.olFormatPlain
        End Sub
    End Class

#End Region

#Region "e-Mail"
    Friend Function SendMail(ByVal SMTPAddress As String, ByRef MailItem As MailItem) As Boolean
        'Dim oApplication As Object
        'Dim oMail As Object
        'Dim oAccount As Object

        'Try
        '    oApplication = Activator.CreateInstance(, "Outlook.Application")
        'Catch ex As Exception
        '    MsgBox("No se ha podido crear una instancia de la aplicación Microsoft Outlook. Verifique que esté instalado en esta computadora.", MsgBoxStyle.Critical, My.Application.Info.Title)
        '    Return False
        'End Try

        'Try
        '    ' Obtengo la cuenta de Outlook a través de la cual se enviará el mail
        '    oAccount = GetAccountForEmailAddress(oApplication, SMTPAddress)
        '    If oAccount Is Nothing Then
        '        Return False
        '    End If
        '    ' Creo el mail
        '    oMail = oApplication.CreateItem(OlItemType.olMailItem)
        '    With oMail
        '        .SendUsingAccount = oAccount
        '        .[To] = MailItem.To
        '        .CC = MailItem.CC
        '        .BCC = MailItem.BCC
        '        .Subject = MailItem.Subject
        '        .BodyFormat = MailItem.BodyFormat
        '        .Body = MailItem.Body
        '        If Not MailItem.RTFBody Is Nothing Then
        '            .RTFBody = MailItem.RTFBody
        '        End If
        '        For Each Attachment As Attachment In MailItem.Attachments
        '            If Attachment.ContentStream Is Nothing Then
        '                .Attachments.Add(Attachment.PathName & Attachment.FileName, Attachment.Type, Attachment.Position, Attachment.DisplayName)
        '            Else
        '                '.Attachments.Add( CreateObject( New createo Attachment.ContentStream, Attachment.DisplayName)
        '            End If
        '        Next
        '        .Send()
        '    End With
        '    Return True

        'Catch ex As Exception
        '    CS_Error.ProcessError(ex, "Error al enviar el e-mail a través de Microsoft Outlook.")
        '    Return False
        'End Try
        Return False
    End Function

    Friend Function GetAccountForEmailAddress(ByVal oApplication As Object, ByVal SMTPAddress As String) As Object
        Dim oAccounts As Object
        Dim oAccount As Object

        Try
            oAccounts = oApplication.Session.Accounts
            For Each oAccount In oAccounts
                If oAccount.SmtpAddress = SMTPAddress Then
                    Return oAccount
                End If
            Next
            MsgBox(String.Format("La cuenta de Microsoft Outlook especificada ({0}) no existe.", SMTPAddress), MsgBoxStyle.Critical, My.Application.Info.Title)
            Return Nothing
        Catch ex As Exception
            CS_Error.ProcessError(ex, "Error al obtener la cuenta de Microsoft Outlook.")
            Return Nothing
        End Try
    End Function
#End Region

End Module
