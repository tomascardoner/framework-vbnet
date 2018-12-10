Imports Microsoft.Office.Interop

Module CS_Office_Outlook_EarlyBinding
    Friend Const SCHEMA_MAILHEADER_FIELD_FROM As String = """urn:schemas:mailheader:from"""
    Friend Const SCHEMA_MAILHEADER_FIELD_TO As String = """urn:schemas:mailheader:to"""
    Friend Const SCHEMA_MAILHEADER_FIELD_SUBJECT As String = """urn:schemas:mailheader:subject"""
    Friend Const SCHEMA_MAILHEADER_FIELD_DATE As String = """urn:schemas:mailheader:date"""

    Friend Const CONTACT_DATE_EMPTYVALUE As Date = #1/1/4501#

    Public WithEvents motkApp As Outlook.Application
    Private mAsynchronousWaiting As Boolean

    '  Friend Function Send(ByVal smtpAddress As String, ByVal Anio As String, ByVal Mes As String, ByVal AttachmentFilename As String) As Boolean
    '      Dim oApp As Outlook.Application
    '      oApp = New Outlook.Application

    '      Dim oMsg As Outlook.MailItem
    '      oMsg = oApp.CreateItem(Outlook.OlItemType.olMailItem)

    '      Dim oAcc As Outlook.Account = GetAccountForEmailAddress(oApp, smtpAddress)
    '      ' Use this account to send the e-mail.
    '      oMsg.SendUsingAccount = oAcc


    '      oMsg.Subject = String.Format("Semanal {0}-{1}", Anio, Mes)
    '      'oMsg.Body = sBody

    '      'oMsg.To = "tomas@cardoner.com.ar"
    '      oMsg.To = "movimientos@minagri.gob.ar"

    '      If AttachmentFilename <> "" Then

    '          oMsg.Attachments.Add(AttachmentFilename)

    '      End If

    '      oMsg.Send()
    '      'MessageBox.Show("Email Send", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '  End Function

    '  Private Function GetAccountForEmailAddress(ByVal application As Interop.Outlook.Application, ByVal smtpAddress As String) As Interop.Outlook.Account
    '      Dim accounts As Interop.Outlook.Accounts = application.Session.Accounts
    '      Dim account As Interop.Outlook.Account

    '' Loop over the Accounts collection of the current Outlook session.
    '      For Each account In accounts
    '          ' When the e-mail address matches, return the account.
    '          If account.SmtpAddress = smtpAddress Then
    '              Return account
    '          End If
    '      Next
    '      Throw New System.Exception(String.Format("No Account with SmtpAddress: {0} exists!", smtpAddress))
    '  End Function

    Friend Function FindRejectedEmails(ByVal FechaDesde As Date, ByVal FechaHasta As Date) As Collection
        Dim otkNameSpace As Outlook.NameSpace
        Dim otkInboxFolder As Outlook.MAPIFolder
        Dim otkSearch As Outlook.Search
        Dim otkResults As Outlook.Results
        Dim otkMessage As Outlook.MailItem

        Dim MessageBody As String
        Dim RejectedAddressSearchStartIndex As Integer
        Dim RejectedAddressSearchCurrentIndex As Integer
        Dim RejectedAddress As String
        Dim CRejectedAddresses As New Collection
        Dim FilterString As String = String.Format("{2} >= '{5}' AND {2} <= '{6}'", SCHEMA_MAILHEADER_FIELD_FROM, SCHEMA_MAILHEADER_FIELD_SUBJECT, SCHEMA_MAILHEADER_FIELD_DATE, My.Settings.Email_SMTP_Username, My.Settings.Email_Delivery_Failed_Subject, FechaDesde.Date.ToString, FechaHasta.Date.AddDays(1).AddSeconds(-1).ToString)

        Try
            motkApp = New Outlook.Application
            otkNameSpace = motkApp.GetNamespace("MAPI")
            otkInboxFolder = otkNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
            otkSearch = motkApp.AdvancedSearch("'" & otkInboxFolder.FolderPath & "'", FilterString, False)
            otkResults = otkSearch.Results()

            If Not otkSearch.IsSynchronous Then
                ' La búsqueda se realiza de forma asincrónica, por lo tanto, hay que esperar que la búsqueda finalice
                mAsynchronousWaiting = True
                Do While mAsynchronousWaiting
                    Application.DoEvents()
                Loop
            End If

            For Each otkMessage In otkResults
                Debug.Print(otkMessage.Sender.Address & " || " & otkMessage.Subject)
                If otkMessage.Sender.Address = My.Settings.Email_Delivery_Failed_SenderAddress AndAlso otkMessage.Subject = My.Settings.Email_Delivery_Failed_Subject Then
                    MessageBody = otkMessage.Body
                    If MessageBody.Contains(My.Settings.Email_Delivery_Failed_ErrorText) Then
                        ' Get the start index to search for the rejected address
                        RejectedAddressSearchStartIndex = MessageBody.IndexOf(My.Settings.Email_Delivery_Failed_RejectedAddress_PreviousText)
                        If RejectedAddressSearchStartIndex > -1 Then
                            RejectedAddressSearchStartIndex = RejectedAddressSearchStartIndex + My.Settings.Email_Delivery_Failed_RejectedAddress_PreviousText.Length + 1

                            ' Search for the rejected address
                            RejectedAddressSearchCurrentIndex = RejectedAddressSearchStartIndex
                            RejectedAddress = ""
                            Do While RejectedAddressSearchCurrentIndex <= MessageBody.Length
                                Debug.Print(Asc(MessageBody.ElementAt(RejectedAddressSearchCurrentIndex).ToString) & " || " & MessageBody.ElementAt(RejectedAddressSearchCurrentIndex).ToString)
                                Select Case MessageBody.ElementAt(RejectedAddressSearchCurrentIndex).ToString
                                    Case Constants.vbLf, Constants.vbCr, Constants.vbNewLine, " "
                                        If RejectedAddress.Length > 0 Then
                                            Exit Do
                                        End If
                                    Case Else
                                        RejectedAddress &= MessageBody.ElementAt(RejectedAddressSearchCurrentIndex).ToString
                                End Select

                                RejectedAddressSearchCurrentIndex += 1
                            Loop
                            If RejectedAddress.Length > 0 Then
                                CRejectedAddresses.Add(RejectedAddress)
                            End If
                        End If
                    End If
                End If
            Next

        Catch ex As Exception
            CS_Error.ProcessError(ex, "Error accediendo a Microsoft Outlook.")
        End Try

        ' Close all
        motkApp = Nothing
        otkNameSpace = Nothing
        otkInboxFolder = Nothing
        otkSearch = Nothing
        otkResults = Nothing
        otkMessage = Nothing

        Return CRejectedAddresses
    End Function

    Public Sub FindRejectedEmails_AdvancedSearchComplete() Handles motkApp.AdvancedSearchComplete
        mAsynchronousWaiting = False
    End Sub

    ''' <summary>
    ''' Establece el valor de la propiedad DisplayName del Email1 en Outlook
    ''' </summary>
    ''' <param name="otkContact"></param>
    ''' <param name="EntidadActual"></param>
    ''' <remarks></remarks>
    Friend Sub Contact_SetDisplayNameOfEmail1(ByRef otkContact As Outlook.ContactItem, ByRef EntidadActual As Entidad)
        If EntidadActual.Email1 Is Nothing Then
            otkContact.Email1DisplayName = Nothing
        Else
            If EntidadActual.Email2 Is Nothing Then
                otkContact.Email1DisplayName = otkContact.LastNameAndFirstName
            Else
                otkContact.Email1DisplayName = otkContact.LastNameAndFirstName & " (" & otkContact.Email1Address & ")"
            End If
        End If
    End Sub

    ''' <summary>
    ''' Establece el valor de la propiedad DisplayName del Email2 en Outlook
    ''' </summary>
    ''' <param name="otkContact"></param>
    ''' <param name="EntidadActual"></param>
    ''' <remarks></remarks>
    Friend Sub Contact_SetDisplayNameOfEmail2(ByRef otkContact As Outlook.ContactItem, ByRef EntidadActual As Entidad)
        If EntidadActual.Email2 Is Nothing Then
            otkContact.Email2DisplayName = Nothing
        Else
            If EntidadActual.Email1 Is Nothing Then
                otkContact.Email2DisplayName = otkContact.LastNameAndFirstName
            Else
                otkContact.Email2DisplayName = otkContact.LastNameAndFirstName & " (" & otkContact.Email1Address & ")"
            End If
        End If
    End Sub
End Module
