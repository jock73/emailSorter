Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports Microsoft.Identity.Client
Imports System.Net
Imports Microsoft.Exchange.WebServices.Data
Imports MailBee.EwsMail
Imports System.Security.Principal
Imports System.Threading
Imports Azure
Imports Newtonsoft.Json.Linq
Imports System.Drawing.Printing

Namespace MailJob




    Class ExchangeClient
        Private ReadOnly _config As MailConfig
        Private _ewsClient As Ews
        Private _newEmail As Dictionary(Of Integer, EwsItem) = New Dictionary(Of Integer, EwsItem)()

        Public ConnectionGood As Boolean = False

        Dim JsonsubjectList As SubjectList


        Public Sub New(config As MailConfig)
            _config = config

            Try
                _ewsClient = New Ews(_config.LicenseKey)
            Catch ex As Exception
                Write("Failed to connect with licensekey:" & ex.Message)

                '' Of course for the Demo, i cant insert real licenskey or 2FA, i have not set up a dummy test account, only live.
                ConnectionGood = False
            End Try

            Do While ConnectionGood



                Dim timeZone = TimeZoneInfo.Local
                Try
                    timeZone = TimeZoneInfo.FindSystemTimeZoneById(config.ExchangeClientTimeZone)
                Catch __unusedInvalidTimeZoneException1__ As InvalidTimeZoneException
                Catch __unusedTimeZoneNotFoundException1__ As TimeZoneNotFoundException
                End Try
                If config.OAuthEnabled Then

                    Write("Connect to Office 365: " & config.ExchangeWebServiceURL)
                    Try
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls
                        Dim auth = ConfidentialClientApplicationBuilder.Create(config.ClientId).WithClientSecret(config.ClientSecret).WithTenantId(config.TenantId).Build()

                        Dim ewsScopes As String() = {config.Scope}
                        Dim authResult = auth.AcquireTokenForClient(ewsScopes).ExecuteAsync().Result
                        _ewsClient.InitEwsClient(ExchangeVersion.Exchange2010_SP1, timeZone)
                        _ewsClient.SetServerUrl(config.ExchangeWebServiceURL)
                        _ewsClient.SetCredentials(New OAuthCredentials(authResult.AccessToken))
                        _ewsClient.Service.ImpersonatedUserId = New ImpersonatedUserId(ConnectingIdType.SmtpAddress, config.ServiceAccountEmail)

                        ConnectionGood = True
                        Write("Connected")
                    Catch ex As Exception
                        Write("Failed to connect:" & ex.Message)

                        ConnectionGood = False
                    End Try



                    JsonsubjectList = New SubjectList(_config.SubjectJsonReplaceFileLocation)


                End If


            Loop
        End Sub


        Public Function Deleteemails(deletesize As Integer) As Integer
            Write("Going to delete items from the deleted box")
            Dim offset = 0




            Dim view = New ItemView(500, offset, OffsetBasePoint.Beginning) With {
                .PropertySet = PropertySet.IdOnly
            }
            view.OrderBy.Add(EmailMessageSchema.DateTimeReceived, SortDirection.Descending)



            Dim findResultsDeleteBox = _ewsClient.DownloadItemsAsync(WellKnownFolderName.DeletedItems, view, False, EwsItemParts.GenericItem Or EwsItemParts.MailMessageBody Or EwsItemParts.MailMessageRecipients Or EwsItemParts.MailMessageFull).Result
            Dim newEmailIndex = 0


            Dim totalEmailCount = findResultsDeleteBox.Count

            Dim middleEmailIndex = totalEmailCount / 2


            Write(totalEmailCount & "Number of Deleted items")

            If totalEmailCount > 200 Then
                Write(totalEmailCount - deletesize & " items to delete")
            End If


            If totalEmailCount > 200 Then

                For i = middleEmailIndex To totalEmailCount - 1
                    Dim item = findResultsDeleteBox(i)
                    _ewsClient.DeleteItemsAsync(item.Id).Wait()
                Next

            End If



        End Function



        Public Function LoadNewEmails(pageSize As Integer) As Integer






            Console.ForegroundColor = ConsoleColor.Blue

            Write("MAX Limit:" & pageSize)

            Dim offset = 0
            Dim view = New ItemView(pageSize, offset, OffsetBasePoint.Beginning) With {
                .PropertySet = PropertySet.IdOnly
            }
            view.OrderBy.Add(EmailMessageSchema.DateTimeReceived, SortDirection.Descending)






            Dim findResults = _ewsClient.DownloadItemsAsync(WellKnownFolderName.Inbox, view, False, EwsItemParts.GenericItem Or EwsItemParts.MailMessageBody Or EwsItemParts.MailMessageRecipients Or EwsItemParts.MailMessageFull).Result
            Dim newEmailIndex = 0




            Write("loading avalible emails:" & findResults.Count.ToString)


            For Each item In findResults
                If item Is Nothing Then
                    Continue For
                End If

                _newEmail.Add(newEmailIndex, item)
                newEmailIndex += 1
                Write("Email Subject:" & item.Subject)

            Next

            If newEmailIndex = 0 Then
                Write("No emails to load, try later")
            End If

            Console.ResetColor()

            Return newEmailIndex
        End Function





        Public Function GetEmailInfo(id As Integer) As EmailMessageInfo
            If Not _newEmail.ContainsKey(id) Then
                Return Nothing
                Write("GetEmailInfo is nothing, exit")
            End If

            Write("eMail no (" & id & ")" & _newEmail(id).Subject)

            Return New EmailMessageInfo() With {
                .Sender = _newEmail(id).From.DisplayName,
                .Subject = _newEmail(id).Subject,
                .myDate = _newEmail(id).DateReceived,
                .Recipient = _newEmail(id).To.AsString,
                .MessageId = _newEmail(id).Id.UniqueId,
                .BCC = _newEmail(id).Bcc.AsString,
                .CC = _newEmail(id).Cc.AsString,
                .SenderEmailAddress = _newEmail(id).From.Email
            }



            Try
                Dim assafsfa As String = _newEmail(id).BodyHtmlText
            Catch ex As Exception

                Dim afasf As String = ""

            End Try


            Try
                Dim assafsfa As String = _newEmail(id).BodyPlainText
            Catch ex As Exception
                Dim afasf As String = ""
            End Try


            Dim asfa As String = ""
        End Function



        Private Function MoveEmail(emailId As Integer) As Task(Of Boolean)

            Try



                Try
                    _ewsClient.DeleteMethod = DeleteMode.MoveToDeletedItems

                    Dim getMailBeeEmailID As ItemId = _newEmail(emailId).UniqueId

                    Dim deletethis As Boolean = _ewsClient.DeleteItemAsync(getMailBeeEmailID).Result



                Catch ex As Exception
                    Write("DeleteEmailFail" & ex.Message)
                End Try




            Catch ex As Exception
                Return Tasks.Task.FromResult(False)
            End Try

            Write("emaildeleted")

            Return Tasks.Task.FromResult(True)



        End Function


        Public Function MoveandDeleteEmail(emailId As Integer) As Boolean
            If Not _newEmail.ContainsKey(emailId) Then
                Return False
            End If


            MoveEmail(emailId)

        End Function


        Private Sub SaveEmailAttachments(message As EwsItem, folderPath As String, info As EmailMessageInfo, ByVal extentionsallowed As String, ByRef messageProcessed As Boolean)

            ''    Console.ForegroundColor = ConsoleColor.Yellow

            messageProcessed = True
            Dim matchedCount = 0
            Dim nullAttachments = 0


            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance)


            Dim countatt As Integer = message.MailBeeMessage.Attachments.Count

            If message.MailBeeMessage.HasAttachments Then
                Dim senderEmailAddressDirectory = String.Empty




                For Each attachment In message.MailBeeMessage.Attachments
                    Dim fileAttachment = TryCast(attachment, MailBee.Mime.Attachment)
                    If fileAttachment IsNot Nothing Then
                        '' If check.Invoke(fileAttachment.Filename) Then
                        matchedCount += 1



                        Dim fileName = Path.Combine(folderPath, fileAttachment.Filename) '' we want to cleanup filenames here.
                        Dim fileNameExtention As String = Path.GetExtension(fileName)
                        Dim extentionallowed = extentionsallowed.Split(",")



                        For i As Integer = 0 To extentionallowed.Length - 1
                            If fileNameExtention.ToLower() = extentionallowed(i).ToLower() Then
                                Write("Attachment:" & fileAttachment.Filename)
                                Dim filenameWithoutExtension = Path.GetFileNameWithoutExtension(fileName)
                                Dim senderFilename = Path.Combine(senderEmailAddressDirectory, filenameWithoutExtension & ".sender")
                                File.WriteAllText(senderFilename, info.SenderEmailAddress)
                                fileAttachment.Save(fileName, True)
                            Else

                                '' no save! 
                            End If
                        Next





                        ''  End If
                    Else
                        ' Donot do anything, this is another email as an attachment
                        nullAttachments += 1
                    End If
                Next

                If message.MailBeeMessage.Attachments.Count > nullAttachments AndAlso matchedCount = 0 Then
                    messageProcessed = False
                End If

            Else

                Write("This email has not attachments")
            End If


            Console.ResetColor()
        End Sub


        Public Sub SaveAttachments(emailId As Integer, folderPath As String, info As EmailMessageInfo, ByVal Exentionsallowed As String, ByRef messageProcessed As Boolean)
            If Not _newEmail.ContainsKey(emailId) Then
                messageProcessed = False
                Return
            End If
            SaveEmailAttachments(_newEmail(emailId), folderPath, info, Exentionsallowed, messageProcessed)

            Thread.Sleep(2000)
        End Sub





        ''' FTP


        Private Sub FTPAttachments(folderPath As String, info As EmailMessageInfo, ByVal extentionsallowed As String, ByRef messageProcessed As Boolean)

            Console.ForegroundColor = ConsoleColor.Green



            Try

                Dim ftpaccount As New FTPconfig()


                Dim filesList As String() = Directory.GetFiles(folderPath)



                For Each filepath In filesList

                    Try
                        Dim filename As String = Path.GetFileName(filepath)

                        Write("From:" & filepath)

                        Dim myFtp As New Utilities.FTP.FTP(ftpaccount)

                        Dim fullpath As String = String.Concat(ftpaccount.ftppath, SubjectManlipulation(info), "/", filename)





                        If myFtp.Upload(filepath, "0", fullpath) Then
                            Write("FTP Sucess to:" & fullpath)
                            messageProcessed = True

                        Else

                            '' if fail, do not process anymore, keep the email in the inbox!

                            Console.ForegroundColor = ConsoleColor.Red
                            Write("FTP Path invalid:" & fullpath)
                            Write("FTP Fail!")
                            Console.ResetColor()

                            messageProcessed = False
                        End If


                    Catch ex As Exception

                        Write("FTP fail" & ex.Message)

                    End Try


                Next




            Catch ex As Exception
                messageProcessed = False
            End Try


            Console.ResetColor()




        End Sub



        Function SubjectManlipulation(info As EmailMessageInfo) As String



            Dim thissubject As String = Trim(SubjectClean(info.Subject))


            For Each item In JsonsubjectList.Subjects '' loop all the email subject settings



                If item.subjectallowwildcard Then

                    If thissubject.ToLower.Contains(item.Subject.ToLower()) Then

                        Return Trim(item.SubjectReplaceTo)
                    End If

                Else

                    If thissubject.ToLower = item.Subject.ToLower Then

                        Return Trim(item.SubjectReplaceTo)

                    End If

                End If




            Next



            Return thissubject

        End Function





        Function SubjectClean(thissubject) As String


            Dim cleanthisSubject As String = thissubject.ToLower()

            Dim cleanabitmore As String = cleanthisSubject


            If _config.CleanSubject.Length > 1 Then

                For Each badcharsinsubject In _config.CleanSubject.Split(",")

                    If cleanabitmore.Contains(badcharsinsubject) Then
                        Write("Cleaning Subject" & badcharsinsubject)
                    End If

                    cleanthisSubject = cleanthisSubject.Replace(badcharsinsubject, "")

                Next badcharsinsubject


            End If




            Return Trim(cleanthisSubject)

        End Function

        Public Sub FTPSavedAttachments(emailId As Integer, folderPath As String, info As EmailMessageInfo, ByVal Exentionsallowed As String, ByRef messageProcessed As Boolean)
            If Not _newEmail.ContainsKey(emailId) Then
                messageProcessed = False
                Return
            End If

            FTPAttachments(folderPath, info, Exentionsallowed, messageProcessed)
        End Sub




    End Class

End Namespace
