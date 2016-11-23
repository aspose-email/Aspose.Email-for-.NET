Imports Aspose.Email.Examples.VisualBasic.Email
Imports Aspose.Email.Examples.VisualBasic.Email.IMAP
Imports Aspose.Email.Examples.VisualBasic.Email.Knowledge.Base
Imports Aspose.Email.Examples.VisualBasic.Email.Outlook
Imports Aspose.Email.Examples.VisualBasic.Email.POP3
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Aspose.Email.Examples.VisualBasic.Email.Exchange
Imports Aspose.Email.Examples.VisualBasic.Email.SMTP
Imports Aspose.Email.Examples.VisualBasic.Email.Thunderbird

Namespace Aspose.Email.Examples.VisualBasic.Email
    Class RunExamples
        Public Shared Sub Main()
            Console.WriteLine("Open RunExamples.vb " & vbLf & "In Main() method uncomment the example that you want to run.")
            Console.WriteLine("=====================================================")

            ' Uncomment the one you want to try out

            ' =====================================================
            ' =====================================================
            ' Email
            ' =====================================================
            ' =====================================================

            'DraftAppointmentRequest.Run()
            'DisplayEmailInformation.Run()
            'ExtractingEmailHeaders.Run()
            'ProcessBouncedMsgs.Run()
            'CreateNewEmail.Run()
            'SaveMessageAsDraft.Run()
            'SpecifyRecipientAddresses.Run()
            'DisplayEmailAddressesNames.Run()
            'SetHTMLBody.Run()
            'SetAlternateText.Run()
            'ManagingEmailAttachments.Run()
            'RemoveAttachments.Run()
            'EmbeddedObjects.Run()
            'LoadMessageWithLoadOptions.Run()
            'SetEmailHeaders.Run()
            'ExtractAttachments.Run()
            'CreateNewMailMessage.Run()
            'ReadMessageByPreservingTNEFAttachments.Run()
            'CreatingTNEFFromMSG.Run()
            'LoadAndSaveFileAsEML.Run()
            'PreserveOriginalBoundaries.Run()
            'PreserveTNEFAttachment.Run()
            'ExtractEmbeddObjects.Run()
            'EncryptAndDecryptMessage.Run()
            'PrintHeaderUsingMhtFormatOptions.Run()
            'ExtraPrintHeaderUsingHideExtraPrintHeader.Run()
            'BayesianSpamAnalyzer.Run()
            'GetDeliveryStatusNotificationMessages.Run()
            'DetectDifferentFileFormats.Run()
            'ExtractEmbeddedObjectsFromEmail.Run()
            'EncryptAndDecryptMessage.Run()
            'AddMapiNoteToPST.Run()
            'ExtractEmbeddObjects.Run()
            'DetectTNEFMessage.Run()
            'DetermineAttachmentEmbeddedMessage.Run()
            'IdentifyInlineAndRegularAttachments.Run()
            'ReceiveNotifications.Run()
            'SetDefaultTextEncoding.Run()
            'SendIMAPasynchronousEmail.Run()


            '/ =====================================================
            '/ =====================================================
            '/ Outlook
            '/ =====================================================
            '/ =====================================================

            'NewPSTAddSubfolders.Run()
            'CreateSaveOutlookFiles.Run()
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
            'DeleteMessagesFromPSTFiles.Run()
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
            'WeeklyEndAfterNoccurrences.Run()
            'EndAfterNoccurrenceSelectMultipleDaysInweek.Run()
            'MonthlyEndAfterNoccurrences.Run()
            'YearlyEndAfterNoccurrences.Run()
            'GenerateRecurrenceFromRecurrenceRule.Run()
            'Exposeproperties.Run()
            'GetTheTextAndRTFBodies.Run()
            'ParseOutlookMessageFile.Run()
            'CreateMeetingRequestWithRecurrence.Run()
            'CreateNewMapiCalendarAndAddToCalendarSubfolder.Run()
            'ConvertMSGToMIMEMessage.Run()

            ' Working with Outlook Personal Storage (PST) files

            'SplitSinglePSTInToMultiplePST.Run()
            'MergeMultiplePSTsInToSinglePST.Run()
            'MergeFolderFromAnotherPSTFile.Run()
            'ConvertOSTToPST.Run()
            'ExtractNumberOfMessages.Run()
            'ExtractAttachmentsFromPSTMessages.Run()
            'AddMessagesToPSTFiles.Run()
            'ReadandConvertOSTFiles.Run()
            'SaveCalendarItems.Run()
            'RetreiveParentFolderInformationFromMessageInfo.Run()
            'ParseSearchableFolders.Run()
            'AccessContactInformation.Run()
            'GetMessageInformation.Run()
            'ChangeFolderContainerClass.Run()
            'CheckPasswordProtection.Run()
            'SetPasswordOnPST.Run()
            'CreateNewPSTFileAndAddingSubfolders.Run()
            'CreateNewMapiContactAndAddToContactsSubfolder.Run()
            'ExtractMessagesFromPSTFile.Run()
            'RemovingPaswordProperty.Run()
            'AddMapiTaskToPST.Run()
            'CreateNewMapiJournalAndAddToSubfolder.Run()
            'AddAttachmentsToMapiJournal.Run()
            'AddMapiCalendarToPST.Run()
            'CreateDistributionListInPST.Run()
            'SaveMessagesDirectlyFromPSTToStream.Run()
            'SearchStringInPSTWithIgnoreCaseParameter.Run()
            'AddFilesToPST.Run()
            'SearchMessagesAndFoldersInPST.Run()
            'MoveItemsToOtherFolders.Run()
            'AddMapiNoteToPST.Run()
            'UpdatePSTCustomProperites.Run()
            'SaveContactInformation.Run()
            'DisplayInformationOfPSTFile.Run()
            'SpecificCriterionSplitPST.Run()


            '/ =====================================================
            '/ =====================================================
            '/ Knowledge-Base
            '/ =====================================================
            '/ =====================================================

            ' PrintEmail.Run()

            '/ =====================================================
            '/ =====================================================
            '/ Exchange
            '/ =====================================================
            '/ =====================================================

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
            'SendCalendarInvitation.Run()

            '/ =====================================================
            '/ =====================================================
            '/ POP3
            '/ =====================================================
            '/ =====================================================

            'ParseMessageAndSave.Run()
            'RecipientInformation.Run()
            'RetrievingEmailHeaders.Run()
            'RetrievingEmailMessages.Run()
            'SaveToDiskWithoutParsing.Run()
            'ConnectingToPOP3.Run()
            'GettingMailboxInfo.Run()
            'SSLEnabledPOP3Server.Run()
            'FilterMessagesFromPOP3Mailbox.Run()
            'RetrieveEmailViaPop3ClientProxyServer.Run()
            'GetServerExtensionsUsingPop3Client.Run()
            'RetrievMessagesAsynchronously.Run()
            'RetrieveMessageSummaryInformationUsingUniqueId.Run()


            '/ =====================================================
            '/ =====================================================
            '/ IMAP
            '/ =====================================================
            '/ =====================================================

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
            'SavingMessagesFromIMAPServer.Run()
            'ListMessagesWithMaximumNumberOfMessages.Run()
            'ListingMessagesRecursively.Run()
            'GetMessageIdUsingImapMessageInfo.Run()
            'FilteringMessagesFromIMAPMailbox.Run()
            'InternalDateFilter.Run()
            'ProxyServerAccessMailbox.Run()
            'RetrievingMessagesAsynchronously.Run()
            'RetreivingServerExtensions.Run()
            'SupportIMAPIdleCommand.Run()

            '/ =====================================================
            '/ =====================================================
            '/ SMTP
            '/ =====================================================
            '/ =====================================================

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
            'SendEmailUsingSMTP.Run()
            'SendEmailAsynchronously.Run()
            'SendingBulkEmails.Run()
            'SendMessageAsTNEF.Run()
            'SendEmailViaProxyServer.Run()
            'SendPlainTextEmailMessage.Run()
            'SendEmailWithAlternateText.Run()
            'ForwardEmail.Run()
            'SignEmailsWithDKIM.Run()

            '/ =====================================================
            '/ =====================================================
            '/ Thunderbird
            '/ =====================================================
            '/ =====================================================

            'ReadMessagesFromThunderbird.Run()
            'CreateNewMessagesToThunderbird.Run()

            ' Stop before exiting
            Console.WriteLine(Environment.NewLine + "Program Finished. Press any key to exit....")
            Console.ReadKey()
        End Sub

        Friend Shared Function GetDataDir_KnowledgeBase() As String
            Return Path.GetFullPath(GetDataDir_Data() + "KnowledgeBase/")
        End Function

        Friend Shared Function Thunderbird() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Thunderbird/"))
        End Function

        Friend Shared Function GetDataDir_Email() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Email/"))
        End Function


        Friend Shared Function GetDataDir_Exchange() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("/Exchange/"))
        End Function

        Friend Shared Function GetDataDir_Outlook() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("Outlook/"))
        End Function

        Friend Shared Function GetDataDir_POP3() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("POP3/"))
        End Function

        Friend Shared Function GetDataDir_IMAP() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("IMAP/"))
        End Function

        Friend Shared Function GetDataDir_SMTP() As String
            Return Path.GetFullPath(GetDataDir_Data() & Convert.ToString("SMTP/"))
        End Function

        Private Shared Function GetDataDir_Data() As String
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
    End Class
End Namespace
