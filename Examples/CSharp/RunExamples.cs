using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Examples.CSharp.Exchange;
using Aspose.Email.Examples.CSharp.Email.IMAP;
using Aspose.Email.Examples.CSharp.Email.Knowledge.Base;
using Aspose.Email.Examples.CSharp.Email.Outlook;
using Aspose.Email.Examples.CSharp.Email.POP3;
using Aspose.Email.Examples.CSharp.Email.IMAP;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Email.Examples.CSharp.Email.Exchange;

namespace Aspose.Email.Examples.CSharp.Email
{
    class RunExamples
    {
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

            // DraftAppointmentRequest.Run();
            // DisplayEmailInformation.Run();
            // ExtractingEmailHeaders.Run();
            // ProcessBouncedMsgs.Run();
            // CreateNewEmail.Run();
            // SaveMessageAsDraft.Run();


            //// =====================================================
            //// =====================================================
            //// Outlook
            //// =====================================================
            //// =====================================================

            //NewPSTAddSubfolders.Run();
            //MergePSTFiles.Run();
            //SplitPST.Run();
            //CreateSaveOutlookFiles.Run();
            //DeleteMessagesFromPSTFiles.Run();                        
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
            //CreatAndSaveCalendaritems.Run();
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
       
            // Working with Outlook Personal Storage (PST) files

            //RetreiveParentFolderInformationFromMessageInfo.Run();
            //ParseSearchableFolders.Run();
            //AccessContactInformation.Run();
            //SaveContactInformation.Run();
            //SaveCalendarItems.Run();
            //AddMessagesToPSTFiles.Run();            
            //DisplayInformationOfPSTFile.Run();
            //ReadandConvertOSTFiles.Run();
            //ConvertOSTToPST.Run();
            //GetMessageInformation.Run();
            //ExtractMessagesFromPSTFile.Run();
            //SaveMessagesDirectlyFromPSTToStream.Run();
            //ExtractNumberOfMessages.Run();
            //CreateNewPSTFileAndAddingSubfolders.Run();
            //ChangeFolderContainerClass.Run();
            //CheckPasswordProtection.Run();
            //RemovingPaswordProperty.Run();
            //SetPasswordOnPST.Run();
            //CreateNewMapiContactAndAddToContactsSubfolder.Run();
            //AddMapiTaskToPST.Run();
            //CreateNewMapiJournalAndAddToSubfolder.Run();
            //AddAttachmentsToMapiJournal.Run();
            //AddMapiCalendarToPST.Run();
            //CreateDistributionListInPST.Run();
            //SearchMessagesAndFoldersInPST.Run();
            //SearchStringInPSTWithIgnoreCaseParameter.Run();
            //MoveItemsToOtherFolders.Run();
            //AddFilesToPST.Run();
            //ExtractAttachmentsFromPSTMessages.Run();


           

            //// =====================================================
            //// =====================================================
            //// Knowledge-Base
            //// =====================================================
            //// =====================================================

            // PrintEmail.Run();

            //// =====================================================
            //// =====================================================
            //// Exchange
            //// =====================================================
            //// =====================================================

            // SecondaryCalendarEvents.Run();          
            // SaveExchangeTaskToDisc.Run();
            // CreateExchangeTask.Run();
            // DeleteExchangeTask.Run();
            // SendExchangeTask.Run();
            // UpdateExchangeTask.Run();

            //// =====================================================
            //// =====================================================
            //// POP3
            //// =====================================================
            //// =====================================================

            // ParseMessageAndSave.Run();
            // RecipientInformation.Run();
            // RetrievingEmailHeaders.Run();
            // RetrievingEmailMessages.Run();
            // SaveToDiskWithoutParsing.Run();
            // ConnectingToPOP3.Run();
            // GettingMailboxInfo.Run();
            // SSLEnabledPOP3Server.Run();

            //// =====================================================
            //// =====================================================
            //// IMAP
            //// =====================================================
            //// =====================================================

            // InsertHeaderAtSpecificLocation.Run();
            // DeletingFolders.Run();
            // RenamingFolders.Run();
            // AddingNewMessage.Run();
            // ConnectingWithIMAPServer.Run();
            // GettingFoldersInformation.Run();
            // MessagesFromIMAPServerToDisk.Run();
            // RemovingMessageFlags.Run();
            // ReadMessagesRecursively.Run();
            // SettingMessageFlags.Run();
            // SSLEnabledIMAPServer.Run();
            // IMAP4IDExtensionSupport.Run();
            // IMAP4ExtendedListCommand.Run();
            // CopyMultipleMessagesFromOneFoldertoAnother.Run();            
            // DeleteSingleMessage.Run();
            // DeleteMultipleMessages.Run();

            //// =====================================================
            //// =====================================================
            //// SMTP
            //// =====================================================
            //// =====================================================

            // SetSpecificIpAddress.Run();
            // ExportAsEML.Run();
            // ImportEML.Run();
            // CustomizingEmailHeader.Run();
            // DeliveryNotifications.Run();
            // SetEmailInfo.Run();
            // SettingHTMLBody.Run();
            // SettingTextBody.Run();
            // AppointmentInICSFormat.Run();
            // CustomizingEmailHeaders.Run();
            // EmbeddedObjects.Run();
            // LoadSmtpConfigFile.Run();
            // MailMerge.Run();
            // ManagingEmailAttachments.Run();
            // MeetingRequests.Run();
            // MultipleRecipients.Run();
            // SendingEMLFilesWithSMTP.Run();
            // SSLEnabledSMTPServer.Run();

            // Stop before exiting
            Console.WriteLine(Environment.NewLine + "Program Finished. Press any key to exit....");
            Console.ReadKey();
        }

        internal static string GetDataDir_KnowledgeBase()
        {
            return Path.GetFullPath("../../Knowledge-Base/Data/");
        }

        internal static string GetDataDir_Email()
        {
            return Path.GetFullPath("../../Email/Data/");
        }

        internal static string GetDataDir_Exchange()
        {
            return Path.GetFullPath("../../Exchange/Data/");
        }

        internal static string GetDataDir_Outlook()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Outlook/");
        }


        internal static string GetDataDir_POP3()
        {
            return Path.GetFullPath("../../POP3/Data/");
        }

        internal static string GetDataDir_IMAP()
        {
            return Path.GetFullPath("../../IMAP/Data/");
        }

        internal static string GetDataDir_SMTP()
        {
            return Path.GetFullPath("../../SMTP/Data/");
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
