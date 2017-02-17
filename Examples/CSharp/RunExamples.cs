using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Examples.CSharp.Exchange;
using Aspose.Email.Examples.CSharp.Email.IMAP;
using Aspose.Email.Examples.CSharp.Email.Knowledge.Base;
using Aspose.Email.Examples.CSharp.Email.Outlook;
using Aspose.Email.Examples.CSharp.Email.POP3;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email.Examples.CSharp.Email.Exchange;
using Aspose.Email.Examples.CSharp.Email.SMTP;
using System.Threading;
using System.Windows.Forms;
using Aspose.Email.Examples.CSharp.Email.Outlook.PST;
using Aspose.Email.Examples.CSharp.Email.Exchange_EWS;
using Aspose.Email.Examples.CSharp.Email.Exchange_WebDav;
using Aspose.Email.Examples.CSharp.Email.Gmail;

namespace Aspose.Email.Examples.CSharp.Email
{
    class RunExamples
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Open RunExamples.cs. \nIn Main() method uncomment the example that you want to run.");
            Console.WriteLine("=====================================================");

                 
            // Uncomment the one you want to try out

            // =====================================================
            // =====================================================
            // Email
            // =====================================================
            // =====================================================

            //ChangeEmailAddress.Run();
            //DetermineAttachmentEmbeddedMessage.Run();
            //DraftAppointmentRequest.Run();
            //DisplayEmailInformation.Run();
            //ExtractingEmailHeaders.Run();
            //ProcessBouncedMsgs.Run();
            //CreateNewEmail.Run();
            //SaveMessageAsDraft.Run();
            //SpecifyRecipientAddresses.Run();
            //DisplayEmailAddressesNames.Run();
            //SetHTMLBody.Run();
            //SetAlternateText.Run();
            //AddEmailAttachments.Run();
            //RemoveAttachments.Run();
            //EmbeddedObjects.Run();
            //LoadMessageWithLoadOptions.Run();
            //SetEmailHeaders.Run();
            //ExtractAttachments.Run();
            //CreateNewMailMessage.Run();
            //ReadMessageByPreservingTNEFAttachments.Run();
            //CreatingTNEFFromMSG.Run();
            //LoadAndSaveFileAsEML.Run();
            //PreserveOriginalBoundaries.Run();
            //PreserveTNEFAttachment.Run();
            //EncryptAndDecryptMessage.Run();
            //PrintHeaderUsingMhtFormatOptions.Run();
            //ExtraPrintHeaderUsingHideExtraPrintHeader.Run();
            //BayesianSpamAnalyzer.Run();
            //GetDeliveryStatusNotificationMessages.Run();
            //DetectDifferentFileFormats.Run();
            //ExtractEmbeddedObjects.Run();
            //ExtractEmbeddedObjectsFromEmail.Run();
            //EncryptAndDecryptMessage.Run();
            //DetectTNEFMessage.Run();
            //ReceiveNotifications.Run();
            //SetDefaultTextEncoding.Run();
            //IdentifyInlineAndRegularAttachments.Run();
            //PreservingEmbeddedMsgFormat.Run();
            //RenderingCalendarEvents.Run();
            //UseMailMessageFeatures.Run();
            //RequestReadReceipt.Run();
            //SpecifyCustomHeader.Run();
            //DisplayAttachmentFileName.Run();
            //GetDecodedHeaderValues.Run();
            //RetrieveContentDescriptionFromAttachment.Run();
            //ExtractMSGEmbeddedAttachment.Run();
            //LoadingEMLAndSavingToMSG.Run();
            //SavingMSGWithPreservedDates.Run();
            //SaveMailMessageAsMHTML.Run();
            //ConvertMHTMLWithOptionalSettings.Run();
            //ExportEmailToMHTWithCustomTimezone.Run();
            //ExportEmailToEML.Run();
            //SaveMessageAsHTML.Run();
            //SaveMessageAsOFT.Run();
            //CheckMessageForEncryption.Run();
            //UpdateTNEFAttachments.Run();
            //AddNewTNEFAttachments.Run();
            //CreateTNEFEMLFromMSG.Run();
            //DetectMessageIsTNEF.Run();
            //CheckBouncedMessage.Run();
            //Application.Run(new EmailValidations());
            //ValidatingEmails.Run();

            //// =====================================================
            //// =====================================================
            //// Outlook
            //// =====================================================
            //// =====================================================

           
            //NewPSTAddSubfolders.Run();
            //CreateSaveOutlookFiles.Run();
            //DeleteBulkItemsFromPSTFile.Run();
            //UpdateBulkMessagesInPSTFile.Run();
            //LoadMSGFiles.Run();
            //LoadingFromStream.Run();
            //GetMAPIProperty.Run();
            //SetMAPIProperties.Run();
            //ReadNamedMAPIProperties.Run();
            //ReadiNamedMAPIPropertyFromAttachment.Run();
            //ReadingNamedMAPIPropertyFromAttachment.Run();
            //RemovePropertiesFromMSGAndAttachments.Run();
            //ConvertEMLToMSG.Run();
            //CreatEMLFileAndConvertToMSG.Run();
            //ReadAndWritingOutlookTemplateFile.Run();          
            //SetFollowUpflag.Run();
            //SetFollowUpForRecipients.Run();
            //MarkFollowUpFlagAsCompleted.Run();
            //RemoveFollowUpflag.Run();
            //ReadFollowupFlagOptionsForMessage.Run();
            //CreateAndSaveOutlookContact.Run();
            //CreatingAndSavingOutlookTasks.Run();
            //AddReminderInformationToMapiTask.Run();
            //AddAttachmentsToMapiTask.Run();
            //AddRecurrenceToMapiTask.Run();
            //CreatAndSaveAnOutlookNote.Run();
            //ReadMapiNote.Run();
            //ConvertMIMEMessagesFromMSGToEML.Run();
            //ConvertMIMEMessageToEML.Run();
            //SetColorCategories.Run();
            //SetReminderByAddingTags.Run();
            //CreateAndSaveCalendaritems.Run();
            //AddDisplayReminderToACalendar.Run();
            //AddAudioReminderToCalendar.Run();
            //ManageAttachmentsFromCalendarFiles.Run();
            //CreatePollUsingMapiMessage.Run();
            //ReadVotingOptionsFromMapiMessage.Run();
            //AddVotingButtonToExistingMessage.Run();
            //DeleteVotingButtonFromMessage.Run();
            //CreateAndSaveDistributionList.Run();
            //CreatReplyMessage.Run();
            //CreateForwardMessage.Run();
            //EndAfterNoccurrences.Run();
            //WeeklyEndAfterNoccurrences.Run();
            //EndAfterNoccurrenceSelectMultipleDaysInweek.Run();
            //MonthlyEndAfterNoccurrences.Run();
            //YearlyEndAfterNoccurrences.Run();
            //GenerateRecurrenceFromRecurrenceRule.Run();
            //ExposeProperties.Run();
            //GetTheTextAndRTFBodies.Run();
            //CreateNewMapiCalendarAndAddToCalendarSubfolder.Run();
            //ParseOutlookMessageFile.Run();
            //ConvertMSGToMIMEMessage.Run();
            //CreatingAndSavingOutlookMessages.Run();
            //CreateMessagesWithAttachments.Run();
            //Application.Run(new WorkingWithMSGAttachments());
            //CreatingMSGFilesWithRTFBody.Run();
            //SavingMessageInDraftStatus.Run();
            //SetBodyCompression.Run();
            //Application.Run(new DragDropOutlookMessages());
            //ReadingVotingOptions.Run();
            //ReadingOnlyVotingButtons.Run();
            //SetAdditionalMAPIProperties.Run();
            //SaveAttachmentsFromOutlookMSGFile.Run();
            //RemoveAttachmentsFromFile.Run();
            //DestroyAttachment.Run();
            //EmbedMessageAsAttachment.Run();
            //ReadEmbeddedMessageFromAttachment.Run();
            //InsertMSGAttachmentAtSpecificlocation.Run();
            //ReplaceEmbeddedMSGAttachmentContents.Run();
            //LoadingContactFromMSG.Run();
            //LoadingContactFromVCard.Run();
            //LoadingContactFromVCardWithSpecifiedEncoding.Run();
            //ReadingMapiTask.Run();
            //ReadingVToDoTask.Run();
            //SavingTheCalendarItemAsMSG.Run();
            //DisplayRecipientsStatusFromMeetingRequest.Run();
            //CreateMapiCalendarTimeZoneFromStandardTimezone.Run();
            //ReadingDistributionListFromPST.Run();
            //SetDailyOccurrenceCount.Run();
            //SetRecurrenceEveryDay.Run();
            //SetDailyNeverEndRecurrence.Run();
            //SetWeeklyRecurrenceMultipleDaysInWeekWithInterval.Run();
            //SetWeeklyEndAfterDateRecurrence.Run();
            //SetWeeklyNeverEndRecurrence.Run();
            //SetMonthlyEndAfterDateRecurrence.Run();
            //SetMonthlyNeverEndRecurrence.Run();
            //YearlyEndAfterDate.Run();
            //SetYearlyNeverEndRecurrence.Run();
           
            // Working with Outlook Personal Storage (PST) files
            //SplitSinglePSTInToMultiplePST.Run();
            //MergeMultiplePSTsInToSinglePST.Run();
            //MergeFolderFromAnotherPSTFile.Run();
            //ConvertOSTToPST.Run();
            //ExtractNumberOfMessages.Run();
            //ExtractAttachmentsFromPSTMessages.Run();
            //AddMessagesToPSTFiles.Run();
            //ReadandConvertOSTFiles.Run();
            //SaveCalendarItems.Run();
            //RetreiveParentFolderInformationFromMessageInfo.Run();
            //ParseSearchableFolders.Run();
            //AccessContactInformation.Run();
            //GetMessageInformation.Run();
            //ChangeFolderContainerClass.Run();
            //CheckPasswordProtection.Run();
            //SetPasswordOnPST.Run();
            //CreateNewPSTFileAndAddingSubfolders.Run();
            //CreateNewMapiContactAndAddToContactsSubfolder.Run();
            //ExtractMessagesFromPSTFile.Run();
            //RemovingPaswordProperty.Run();
            //AddMapiTaskToPST.Run();
            //CreateNewMapiJournalAndAddToSubfolder.Run();
            //AddAttachmentsToMapiJournal.Run();
            //AddMapiCalendarToPST.Run();
            //CreateDistributionListInPST.Run();
            //SaveMessagesDirectlyFromPSTToStream.Run();
            //SearchStringInPSTWithIgnoreCaseParameter.Run();
            //AddFilesToPST.Run();
            //SearchMessagesAndFoldersInPST.Run();
            //MoveItemsToOtherFolders.Run();
            //AddMapiNoteToPST.Run();
            //UpdatePSTCustomProperites.Run();
            //SaveContactInformation.Run();
            //DisplayInformationOfPSTFile.Run();
            //SpecificCriterionSplitPST.Run();
            //AddMessagesFromOtherPST.Run();
            //DeleteMessagesFromPSTFiles.Run();
            //MergeFolderFromAnotherPSTFile.Run();
            //LoadingPSTFile.Run();

            //// =====================================================
            //// =====================================================
            //// Knowledge-Base
            //// =====================================================
            //// =====================================================

            //PrintEmail.Run();

            //// =====================================================
            //// =====================================================
            //// Exchange
            //// =====================================================
            //// =====================================================

            //GetMailboxInformationFromExchangeWebServices.Run();
            //GetMailboxInformationFromExchangeServer.Run();
            //ListExchangeServerMessages.Run();
            //ListMessagesFromDifferentFolders.Run();
            //ListingMessagesFromFolders.Run();
            //ListMessagesByID.Run();
            //EnumeratMessagesWithPaginginEWS.Run();
            //SaveMessagesFromExchangeServerMailboxToEML.Run();
            //SaveMessagesUsingExchangeWebServices.Run();
            //SaveMessagesToMemoryStream.Run();
            //SaveMessagesToMemoryStreamUsingEWS.Run();
            //ExchangeClientSaveMessagesInEMLFormat.Run();
            //SendEmailMessagesUsingExchangeServer.Run();
            //SendEmailMessagesUsingExchangeWebServices.Run();
            //MoveMessageFromOneFolderToAnotherUsingExchangeClient.Run();
            //MoveMessageFromOneFolderToAnotherusingEWS.Run();
            //DeleteMessagesFromExchangeServer.Run();
            //DeleteMessagesFromusingEWS.Run();
            //DownloadMessagesFromExchangeServerFoldersRecursively.Run();
            //DownloadMessagesFromPublicFolders.Run();
            //ConnectExchangeServerUsingIMAP.Run();
            //ExchangeServerUsesSSL.Run();
            //SendMeetingRequestsUsingExchangeServer.Run();
            //SendMeetingRequestsUsingEWS.Run();
            //FilterMessagesFromExchangeMailbox.Run();
            //ExchangeServerReadRules.Run();
            //CreateNewRuleOntheExchangeServer.Run();
            //UpdateRuleOntheExchangeServer.Run();
            //ReadUserConfiguration.Run();
            //CreatUserConfigurations.Run();
            //UpdateUserConfiguration.Run();
            //DeleteUserConfiguration.Run();
            //FindConversationsOnExchangeServer.Run();
            //CopyConversations.Run();
            //MoveConversations.Run();
            //DeleteConversations.Run();
            //GetMailTips.Run();
            //AccessAnotherMailboxUsingExchangeClient.Run();
            //AccessAnotherMailboxUsingExchangeWebServiceClient.Run();
            //AccessCustomFolderUsingExchangeWebServiceClient.Run();
            //ExchangeImpersonationUsingEWS.Run();
            //RetrieveFolderPermissionsUsingExchangeWebServiceClient.Run();
            //SendTaskRequestUsingIEWSClient.Run();
            //ExchangeFoldersBackupToPST.Run();
            //CreateREAndFWMessages.Run();
            //CreateAndSendingMessageWithVotingOptions.Run();
            //PreFetchMessageSizeUsingIEWSClient.Run();
            //SynchronizeFolderItems.Run();
            //SecondaryCalendarEvents.Run();
            //SaveExchangeTaskToDisc.Run();
            //CreateExchangeTask.Run();
            //DeleteExchangeTask.Run();
            //SendExchangeTask.Run();
            //UpdateExchangeTask.Run();  
            //SendCalendarInvitation.Run();
            //ListTasksFromExchangeServer.Run();
            //AddContactsToExchangeServerUsingEWS.Run();
            //FilterTasksFromServer.Run();
            //ConnectingToExchangeServerUsingEWS.Run();
            //SaveMessagesUsingIMAP.Run();
            //ListingMessagesUsingEWS.Run();
            //SaveMessagesInMSGFormatExchangeEWS.Run();
            //GetExchangeMessageInfoFromMessageURI.Run();
            //FetchMessageUsingEWS.Run();
            //FilterMessagesUsingEWS.Run();
            //FilterMessagesOnCriteriaUsingEWS.Run();
            //FilterWithComplexQueriesUsingEWS.Run();
            //CaseSensitiveEmailsFilteringUsingEWS.Run();
            //CreatingUpdatingAndDeletingCalendarItemsUsingEWS.Run();
            //GettingContactsUsingEWS.Run();
            //ResolveContactsUsingContactName.Run();
            //FetchContactUsingId.Run();
            //UpdateContactInformationUsingEWS.Run();
            //DeleteContactsFromExchangeServerUsingEWS.Run();
            //ExpandPublicDistributionList.Run();
            //SendEmailToPrivateDistributionList.Run();
            //SpecifyTimeZoneForExchange.Run();
            //GettingUnifiedMessagingConfigurationInformation.Run();
            //ConnectingToExchange.Run();
            //SaveMessagesInMSGFormatExchangeClient.Run();
            //FetchMessageUsingExchangeClient.Run();
            //FilterMessagesOnCriteriaUsingExchangeClient.Run();
            //FilterWithComplexQueriesUsingExchangeClient.Run();
            //CaseSensitiveEmailsFilteringUsingExchangeClient.Run();
            //GettingContactsFromAnExchangeServer.Run(); 
            
            //// =====================================================
            //// =====================================================
            //// Gmail
            //// =====================================================
            //// =====================================================

            //InsertFetchAndUpdateCalendar.Run();
            //DeleteParticularCalendar.Run();
            //AccessColorInfo.Run();
            //AccessGmailContacts.Run();
            //CreateGmailContact.Run();
            //UpdateGmailContact.Run();
            //DeleteGmailContact.Run();
            //SavingContact.Run();

            //// =====================================================
            //// =====================================================
            //// POP3
            //// =====================================================
            //// =====================================================

            //ParseMessageAndSave.Run();
            //RecipientInformation.Run();
            //RetrievingEmailHeaders.Run();
            //RetrievingEmailMessages.Run();
            //SaveToDiskWithoutParsing.Run();
            //ConnectingToPOP3.Run();
            //GettingMailboxInfo.Run();
            //SSLEnabledPOP3Server.Run();
            //FilterMessagesFromPOP3Mailbox.Run();
            //RetrieveEmailViaPop3ClientProxyServer.Run();
            //GetServerExtensionsUsingPop3Client.Run();
            //RetrieveMessagesAsynchronously.Run();
            //RetrieveMessageSummaryInformationUsingUniqueId.Run();
            //DeleteEmailByIndex.Run();
            //DeleteAllEmails.Run();
            //CancelDeletes.Run();
            //Pop3ClientActivityLogging.Run();
            //GetEmailCountIntheMailbox.Run();
            //GetMessagesUsingSpecificCriteria.Run();
            //BuildComplexQueries.Run();
            //ApplyCaseSensitiveFilters.Run();
            //ListMessagesAsynchronouslyWithMailQuery.Run();

            //// =====================================================
            //// =====================================================
            //// IMAP
            //// =====================================================
            //// =====================================================

            //InsertHeaderAtSpecificLocation.Run();
            //DeletingFolders.Run();
            //RenamingFolders.Run();
            //AddingNewMessage.Run();
            //ConnectingWithIMAPServer.Run();
            //GettingFoldersInformation.Run();
            //MessagesFromIMAPServerToDisk.Run();
            //RemovingMessageFlags.Run();
            //ReadMessagesRecursively.Run();
            //SettingMessageFlags.Run();
            //SSLEnabledIMAPServer.Run();
            //IMAP4IDExtensionSupport.Run();
            //IMAP4ExtendedListCommand.Run();
            //CopyMultipleMessagesFromOneFoldertoAnother.Run();
            //DeleteSingleMessage.Run();
            //DeleteMultipleMessages.Run();
            //SavingMessagesFromIMAPServer.Run();
            //ListMessagesWithMaximumNumberOfMessages.Run();
            //ListingMessagesRecursively.Run();
            //GetMessageIdUsingImapMessageInfo.Run();
            //FilteringMessagesFromIMAPMailbox.Run();
            //InternalDateFilter.Run();
            //AccessMailboxViaProxyServer.Run();
            //RetrievingMessagesAsynchronously.Run();
            //RetreivingServerExtensions.Run();
            //SupportIMAPIdleCommand.Run();
            //ListMessagesAsynchronously.Run();
            //ListingMIMEMessageIdInImapMessageInfo.Run();
            //GetMessagesWithSpecificCriteria.Run();
            //BuildingComplexQueries.Run();
            //CaseSensitiveEmailsFiltering.Run();
            //SpecifyEncodingForQueryBuilder.Run();
            
            //// =====================================================
            //// =====================================================
            //// SMTP
            //// =====================================================
            //// =====================================================

            //SetSpecificIpAddress.Run();
            //ExportAsEML.Run();
            //ImportEML.Run();
            //CustomizingEmailHeader.Run();
            //DeliveryNotifications.Run();
            //SetEmailInfo.Run();
            //SettingHTMLBody.Run();
            //SettingTextBody.Run();
            //AppointmentInICSFormat.Run();
            //CustomizingEmailHeaders.Run();
            //EmbeddedObjects.Run();
            //LoadSmtpConfigFile.Run();
            //MailMerge.Run();
            //MeetingRequests.Run();
            //MultipleRecipients.Run();
            //SendingEMLFilesWithSMTP.Run();
            //SSLEnabledSMTPServer.Run();
            //SendEmailUsingSMTP.Run();
            //SendEmailAsynchronously.Run();
            //SendingBulkEmails.Run();
            //SendMessageAsTNEF.Run();
            //SendEmailViaProxyServer.Run();
            //SendPlainTextEmailMessage.Run();
            //SendEmailWithAlternateText.Run();
            //ForwardEmail.Run();
            //SignEmailsWithDKIM.Run();
            //CreateMeetingRequestWithRecurrence.Run();
            //LoadingEMLFilesFromDisk.Run();
            //SendEmailsSynchronously.Run();
            //UseSmtpClientFeatures.Run();
            //SendingEmailWithAlternateText.Run();
            //RetreiveServerExtensions.Run();
            //SMTPClientActivityLogging.Run();
            //UsingDetachedCertificate.Run();

            //// =====================================================
            //// =====================================================
            //// Thunderbird
            //// =====================================================
            //// =====================================================

            //ReadMessagesFromThunderbird.Run();
            //CreateNewMessagesToThunderbird.Run();

            // Stop before exiting
            
            Console.WriteLine(Environment.NewLine + "Program Finished. Press any key to exit....");
            Console.ReadKey();
        }

        internal static string GetDataDir_KnowledgeBase()
        {
            return Path.GetFullPath(GetDataDir_Data() + "KnowledgeBase/");
        }

        internal static string Thunderbird()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Thunderbird/");
        }

        internal static string GetDataDir_Email()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Email/");
        }

        internal static string GetDataDir_Exchange()
        {
            return Path.GetFullPath(GetDataDir_Data() + "/Exchange/");
        }

        internal static string GetDataDir_Outlook()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Outlook/");
        }

        internal static string GetDataDir_POP3()
        {
            return Path.GetFullPath(GetDataDir_Data() + "POP3/");
        }

        internal static string GetDataDir_IMAP()
        {
            return Path.GetFullPath(GetDataDir_Data() + "IMAP");
        }

        internal static string GetDataDir_SMTP()
        {
            return Path.GetFullPath(GetDataDir_Data() + "SMTP/");
        }

        internal static string GetDataDir_Gmail()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Gmail/");
        }

        private static string GetDataDir_Data()
        {
            var parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
            string startDirectory = null;
            if (parent != null)
            {
                var directoryInfo = parent.Parent;
                if (directoryInfo != null)
                {
                    startDirectory = directoryInfo.FullName;
                }
            }
            else
            {
                startDirectory = parent.FullName;
            }
            return startDirectory != null ? Path.Combine(startDirectory, "Data\\") : null;
        }
    }
}