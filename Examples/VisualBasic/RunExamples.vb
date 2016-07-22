Imports System.IO
Imports Aspose.Email.Examples.VisualBasic.Email
Imports Aspose.Email.Examples.VisualBasic.Email.Exchange
Imports Aspose.Email.Examples.VisualBasic.Email.Knowledge.Base
Imports Aspose.Email.Examples.VisualBasic.Email.Outlook
Imports Aspose.Email.Examples.VisualBasic.Email.POP3
Imports Aspose.Email.Examples.VisualBasic.Email.SMTP
Imports Aspose.Email.Examples.VisualBasic.Email.IMAP
Imports Aspose.Email.Examples.VisualBasic.Email.Personal.Outlook

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

            'DraftAppointmentRequest.Run()
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
            'DeleteMessagesFromPSTFiles.Run()
            'DeleteBulkItemsFromPSTFile.Run()
            'UpdateBulkMessagesInPSTFile.Run()

            'LoadMSGFiles.Run()
            'LoadingFromStream.Run()
            'GetMAPIProperty.Run()
            'SetMAPIProperties.Run()
            'ReadNamedMAPIProperties.Run()
            'ReadiNamedMAPIPropertyFromAttachment.Run()
            'ReadingNamedMAPIPropertyFromAttachment.Run()
            'RemovePropertiesFromMSGAndAttachments.Run()
            'ConvertEMLToMSG.Run()
            'CreatEMLFileAndConvertToMSG.Run()
            'ReadAndWritingOutlookTemplateFile.Run()
            'SetFollowUpflag.Run()
            'SetFollowUpForRecipients.Run()
            'MarkFollowUpFlagAsCompleted.Run()
            'RemoveFollowUpflag.Run()
            'ReadFollowupFlagOptionsForMessage.Run()
            'CreateAndSaveOutlookContact.Run()
            'CreatingAndSavingOutlookTasks.Run()
            'AddReminderInformationToMapiTask.Run()
            'AddAttachmentsToMapiTask.Run()
            'AddRecurrenceToMapiTask.Run()
            'CreatAndSaveAnOutlookNote.Run()
            'ReadMapiNote.Run()
            'ConvertMIMEMessagesFromMSGToEML.Run()
            'ConvertMIMEMessageToEML.Run()
            'SetColorCategories.Run()
            'SetReminderByAddingTags.Run()
            'CreatAndSaveCalendaritems.Run()
            'AddDisplayReminderToACalendar.Run()
            'AddAudioReminderToCalendar.Run()
            'ManageAttachmentsFromCalendarFiles.Run()
            'CreatePollUsingMapiMessage.Run()
            'ReadVotingOptionsFromMapiMessage.Run()
            'AddVotingButtonToExistingMessage.Run()
            'DeleteVotingButtonFromMessage.Run()
            'CreateAndSaveDistributionList.Run()
            'CreatReplyMessage.Run()
            'CreateForwardMessage.Run()
            'EndAfterNoccurrences.Run()
            'GenerateRecurrenceFromRecurrenceRule.Run()
            'RetreiveParentFolderInformationFromMessageInfo.Run()
            'ParseSearchableFolders.Run()
            'AccessContactInformation.Run()
            'SaveContactInformation.Run()
            'SaveCalendarItems.Run()
            'AddMessagesToPSTFiles.Run()
            'DisplayInformationOfPSTFile.Run()
            'ReadandConvertOSTFiles.Run()
            'ConvertOSTToPST.Run()
            'GetMessageInformation.Run()
            'ExtractMessagesFromPSTFile.Run()
            'SaveMessagesDirectlyFromPSTToStream.Run()
            'ExtractNumberOfMessages.Run()
            'CreateNewPSTFileAndAddingSubfolders.Run()
            'ChangeFolderContainerClass.Run()
            'CheckPasswordProtection.Run()
            'RemovingPaswordProperty.Run()
            'SetPasswordOnPST.Run()
            'CreateNewMapiContactAndAddToContactsSubfolder.Run()
            'AddMapiTaskToPST.Run()
            'CreateNewMapiJournalAndAddToSubfolder.Run()
            'AddAttachmentsToMapiJournal.Run()
            'AddMapiCalendarToPST.Run()
            'CreateDistributionListInPST.Run()
            'SearchMessagesAndFoldersInPST.Run()
            'SearchStringInPSTWithIgnoreCaseParameter.Run()
            'MoveItemsToOtherFolders.Run()
            'AddFilesToPST.Run()
            'ExtractAttachmentsFromPSTMessages.Run()

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

            'GetMailboxInformationFromExchangeWebServices.Run()
            'GetMailboxInformationFromExchangeServer.Run()
            'GetMailboxInformationFromExchangeServer.Run()
            'ListExchangeServerMessages.Run()
            'ProgrammingSamplesUsingEWS.Run()
            'ListMessagesFromDifferentFolders.Run()
            'ExchangeServerUsingEWS.Run()
            'ListMessagesByID.Run()
            'EnumeratMessagesWithPaginginEWS.Run()
            'SaveMessagesFromExchangeServerMailboxToEML.Run()
            'SaveMessagesUsingExchangeWebServices.Run()
            'SaveMessagesToMemoryStream.Run()
            'SaveMessagesToMemoryStreamUsingEWS.Run()
            'ExchangeClientSaveMessagesInMSGFormat.Run()
            'SendEmailMessagesUsingExchangeServer.Run()
            'SendEmailMessagesUsingExchangeWebServices.Run()
            'MoveMessageFromOneFolderToAnotherUsingExchangeClient.Run()
            'MoveMessageFromOneFolderToAnotherusingEWS.Run()
            'DeleteMessagesFromExchangeServer.Run()
            'DeleteMessagesFromusingEWS.Run()
            'DownloadMessagesFromExchangeServerFoldersRecursively.Run()
            'DownloadMessagesFromPublicFolders.Run()
            'ConnectExchangeServerUsingIMAP.Run()
            'ExchangeServerUsesSSL.Run()
            'SendMeetingRequestsUsingExchangeServer.Run()
            'SendMeetingRequestsUsingEWS.Run()
            'FilterMessagesFromExchangeMailbox.Run()
            'ExchangeServerReadRules.Run()
            'CreateNewRuleOntheExchangeServer.Run()
            'UpdateRuleOntheExchangeServer.Run()
            'ReadUserConfiguration.Run()
            'CreateUserConfigurations.Run()
            'UpdateUserConfiguration.Run()
            'DeleteUserConfiguration.Run()
            'FindConversationsOnExchangeServer.Run()
            'CopyConversations.Run()
            'MoveConversations.Run()
            'DeleteConversations.Run()
            'GetMailTips.Run()
            'AccessAnotherMailboxUsingExchangeClient.Run()
            'AccessAnotherMailboxUsingExchangeWebServiceClient.Run()
            'AccessCustomFolderUsingExchangeWebServiceClient.Run()
            'ExchangeImpersonationUsingEWS.Run()
            'RetrieveFolderPermissionsUsingExchangeWebServiceClient.Run()
            'SendTaskRequestUsingIEWSClient.Run()
            'ExchangeFoldersBackupToPST.Run()
            'CreateREAndFWMessages.Run()
            'CreateAndSendingMessageWithVotingOptions.Run()
            'PreFetchMessageSizeUsingIEWSClient.Run()
            'SynchronizeFolderItems.Run()
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
            'IMAP4IDExtensionSupport.Run()
            'IMAP4ExtendedListCommand.Run()
            'CopyMultipleMessagesFromOneFoldertoAnother.Run()
            'DeleteSingleMessage.Run()
            'DeleteMultipleMessages.Run()

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
            Return Path.GetFullPath(GetDataDir_Data() + "Outlook/")
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


        Private Function GetDataDir_Data() As String
            Dim parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent
            Dim startDirectory As String = Nothing
            If parent IsNot Nothing Then
                Dim directoryInfo = parent.Parent
                If directoryInfo IsNot Nothing Then
                    startDirectory = directoryInfo.FullName
                End If
            Else
                startDirectory = parent.FullName
            End If
            Return If(startDirectory IsNot Nothing, Path.Combine(startDirectory, "Data\"), Nothing)
        End Function


    End Module
End Namespace