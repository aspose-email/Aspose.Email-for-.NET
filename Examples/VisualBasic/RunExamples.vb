Imports System.IO
Imports Aspose.Email.Examples.VisualBasic.Email
Imports Aspose.Email.Examples.VisualBasic.Exchange
Imports Aspose.Email.Examples.VisualBasic.IMAP
Imports Aspose.Email.Examples.VisualBasic.Knowledge.Base
Imports Aspose.Email.Examples.VisualBasic.Knowledge.Outlook
Imports Aspose.Email.Examples.VisualBasic.Knowledge.POP3
Imports Aspose.Email.Examples.VisualBasic.Knowledge.SMTP

Namespace Aspose.Email.Examples.VisualBasic
    Module RunExamples

        Sub Main()

            Console.WriteLine("Open RunExamples.vb. " & vbLf & "In Main() method uncomment the example that you want to run.")
            Console.WriteLine("=====================================================")

            ' Uncomment the one you want to try out

            '' =====================================================
            '' =====================================================
            '' Email
            '' =====================================================
            '' =====================================================

            DraftAppointmentRequest.Run()
            'DisplayEmailInformation.Run()
            'ExtractingEmailHeaders.Run()
            'ProcessBouncedMsgs.Run()
            'CreateNewEmail.Run()
            'SaveMessageAsDraft.Run()

            '' =====================================================
            '' =====================================================
            '' Outlook
            '' =====================================================
            '' =====================================================

            'NewPSTAddSubfolders.Run()
            'MergePSTFiles.Run()
            'SplitPST.Run()
            'CreateSaveOutlookFiles.Run()

            '' =====================================================
            '' =====================================================
            '' Knowledge-Base
            '' =====================================================
            '' =====================================================

            'PrintEmail.Run()

            '' =====================================================
            '' =====================================================
            '' Exchange
            '' =====================================================
            '' =====================================================

            'SecondaryCalendarEvents.Run()
            'SaveExchangeTaskToDisc.Run()
            'CreateExchangeTask.Run()
            'DeleteExchangeTask.Run()
            'SendExchangeTask.Run()
            'UpdateExchangeTask.Run()

            '' =====================================================
            '' =====================================================
            '' POP3
            '' =====================================================
            '' =====================================================

            'ParseMessageAndSave.Run()
            'RecipientInformation.Run()
            'RetrievingEmailHeaders.Run()
            'RetrievingEmailMessages.Run()
            'SaveToDiskWithoutParsing.Run()
            'ConnectingToPOP3.Run()
            'GettingMailboxInfo.Run()
            'SSLEnabledPOP3Server.Run()

            '' =====================================================
            '' =====================================================
            '' IMAP
            '' =====================================================
            '' =====================================================

            'InsertHeaderAtSpecificLocation.Run()
            'DeletingFolders.Run()
            'RenamingFolders.Run()
            'AddingNewMessage.Run()
            'ConnectingWithIMAPServer.Run()
            'GettingFoldersInformation.Run()
            'MessagesFromIMAPServerToDisk.Run()
            'RemovingMessageFlags.Run()
            'ReadMessagesRecursively.Run()
            'SettingMessageFlags.Run()
            'SSLEnabledIMAPServer.Run()

            '' =====================================================
            '' =====================================================
            '' SMTP
            '' =====================================================
            '' =====================================================


            'SetSpecificIpAddress.Run()
            'ExportAsEML.Run()
            'ImportEML.Run()
            'CustomizingEmailHeader.Run()
            'DeliveryNotifications.Run()
            'SetEmailInfo.Run()
            'SettingHTMLBody.Run()
            'SettingTextBody.Run()
            'AppointmentInICSFormat.Run()
            'CustomizingEmailHeaders.Run()
            'EmbeddedObjects.Run()
            'LoadSmtpConfigFile.Run()
            'MailMerge.Run()
            'ManagingEmailAttachments.Run()
            'MeetingRequests.Run()
            'MultipleRecipients.Run()
            'SendingEMLFilesWithSMTP.Run()
            'SSLEnabledSMTPServer.Run()

            ' Stop before exiting
            Console.WriteLine(Environment.NewLine + "Program Finished. Press any key to exit....")
            Console.ReadKey()
        End Sub

        Friend Function GetDataDir_KnowledgeBase() As String
            Return Path.GetFullPath("../../Knowledge-Base/Data/")
        End Function

        Friend Function GetDataDir_Email() As String
            Return Path.GetFullPath("../../Email/Data/")
        End Function

        Friend Function GetDataDir_Exchange() As String
            Return Path.GetFullPath("../../Exchange/Data/")
        End Function

        Friend Function GetDataDir_Outlook() As String
            Return Path.GetFullPath("../../Outlook/Data/")
        End Function

        Friend Function GetDataDir_POP3() As String
            Return Path.GetFullPath("../../POP3/Data/")
        End Function

        Friend Function GetDataDir_IMAP() As String
            Return Path.GetFullPath("../../IMAP/Data/")
        End Function

        Friend Function GetDataDir_SMTP() As String
            Return Path.GetFullPath("../../SMTP/Data/")
        End Function

    End Module
End Namespace