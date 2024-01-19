Imports System
Imports MailBee
Imports MailBee.Mime
Imports MailBee.EwsMail
Imports Microsoft.Exchange.WebServices.Data
Imports Microsoft.Identity.Client
Imports System.Net
Imports System.Reflection.Metadata.Ecma335
Imports System.Threading
Imports System.Diagnostics.Metrics
Imports System.IO
Imports MailBee.ImapMail
Imports PareX.MailJob
Imports Google.Apis
Imports System.Net.WebRequestMethods
Imports System.Globalization

Module Program


    Dim CountEmails As Integer = 0
    Dim ftpaccount As New FTPconfig()

    Dim messageProceed As Boolean = True
    Dim config As New MailConfig()


    Sub Main(args As String())

        Console.ForegroundColor = ConsoleColor.Gray



        logfile()


        Write("The server .NET framework version is: " & System.Environment.Version.ToString())





        Dim Office365 As New ExchangeClient(config)

        If Not Office365.ConnectionGood Then
            Exit Sub '' send error
        End If

        ''    Dim deletesomeOldemails = Office365.Deleteemails(50) new feauture

        CountEmails = Office365.LoadNewEmails(config.MaxEmails)

        messageProceed = DeleteAllAtachments(config.SaveFileLocalLocation, "PrePurge:")

        For i As Int16 = 0 To CountEmails

            Dim tempInfo As EmailMessageInfo = Office365.GetEmailInfo(i)

            Office365.SaveAttachments(i, config.SaveFileLocalLocation, tempInfo, ftpaccount.ExtentionsAllowed, messageProceed)

            Office365.FTPSavedAttachments(i, config.SaveFileLocalLocation, tempInfo, ftpaccount.ExtentionsAllowed, messageProceed)

            If messageProceed Then
                Office365.MoveandDeleteEmail(i)
            End If

            messageProceed = DeleteAllAtachments(config.SaveFileLocalLocation, "PostPurge")


            Write(tempInfo.Subject)
        Next






        Write("Email Sorter End" & DateTime.Now)

        Write("Pause 10 seconds")
        Thread.Sleep(10000)



        If config.PauseConsole.ToString.ToLower() = "false" Then
            Console.Clear()
        End If


        Close()



    End Sub




    Private Function DeleteAllAtachments(ByVal directoryPath As String, ByVal PrePost As String) As Boolean


        ''''''' CREATE DIR IF NEEDED
        If Not Directory.Exists(directoryPath) Then
            Directory.CreateDirectory(directoryPath)
            Write(PrePost & " Directory created: " & directoryPath)
        End If


        ''''''  GET LIST IF FILES IN DIR
        Dim files As String() = Nothing
        Try
            files = Directory.GetFiles(directoryPath)
        Catch ex As Exception
            Write(PrePost & " DeleteAllAtachments Error:" & ex.Message)
            Return False
        End Try

        If files.Length > 0 Then
            Write(PrePost & " local files(" & files.Length & ") to purge")
        Else
            Write(PrePost & " local files nothing to purge")
            Return True
        End If



        Try

            For Each file As String In files
                System.IO.File.Delete(file)
            Next

        Catch ex As Exception
            Write(PrePost & " fail " & ex.Message)
            Return False
        End Try


        Dim validatePurge As String() = Directory.GetFiles(directoryPath)


        If validatePurge.Length = 0 Then
            Write(PrePost & "Sucess")
            Return True
        Else
            Write(PrePost & "fail")
            Return False
        End If


    End Function

    Public Function getFTPaccount()




        Dim myFtp As New Utilities.FTP.FTP(ftpaccount)

        Return myFtp


        ''  Dim suc As Boolean = myFtp.Upload(x.SaveAttachmentLocation + Attachment.SafeFileName, "0", x.FTPEdixPath & "\" & subformat & "\" & Attachment.SafeFileName)



        Return True
    End Function



End Module



