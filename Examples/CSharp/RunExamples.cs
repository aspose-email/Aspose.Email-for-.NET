using Aspose.Email.Examples.CSharp.Email;
using Aspose.Email.Examples.CSharp.Exchange;
using Aspose.Email.Examples.CSharp.IMAP;
using Aspose.Email.Examples.CSharp.Knowledge.Base;
using Aspose.Email.Examples.CSharp.Outlook;
using Aspose.Email.Examples.CSharp.POP3;
using Aspose.Email.Examples.CSharp.SMTP;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Aspose.Email.Examples.CSharp
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

            DraftAppointmentRequest.Run();
            //DisplayEmailInformation.Run();
            //ExtractingEmailHeaders.Run();
            //ProcessBouncedMsgs.Run();
            //CreateNewEmail.Run();
            //SaveMessageAsDraft.Run();

            //// =====================================================
            //// =====================================================
            //// Outlook
            //// =====================================================
            //// =====================================================

            //NewPSTAddSubfolders.Run();
            //MergePSTFiles.Run();
            //SplitPST.Run();
            //CreateSaveOutlookFiles.Run();

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

            //SecondaryCalendarEvents.Run();          
            //SaveExchangeTaskToDisc.Run();
            //CreateExchangeTask.Run();
            //DeleteExchangeTask.Run();
            //SendExchangeTask.Run();
            //UpdateExchangeTask.Run();

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
            //ManagingEmailAttachments.Run();
            //MeetingRequests.Run();
            //MultipleRecipients.Run();
            //SendingEMLFilesWithSMTP.Run();
            //SSLEnabledSMTPServer.Run();

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
            return Path.GetFullPath("../../Outlook/Data/");
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
    }
}
