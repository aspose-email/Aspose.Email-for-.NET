/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. 
   Please check https://docs.nuget.org/consume/nuget-faq for more information. 
   If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. 
   For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx  
*/

using CSharp.Email;
using CSharp.Exchange;
using CSharp.IMAP;
using CSharp.Knowledge_Base;
using CSharp.Outlook;
using CSharp.POP3;
using CSharp.SMTP;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp
{
    class RunExamples
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Open RunExamples.cs. In Main() method, Un-comment the example that you want to run");
            Console.WriteLine("=====================================================");
            // Un-comment the one you want to try out

            // =====================================================
            // =====================================================
            // Email
            // =====================================================
            // =====================================================

            //DraftAppointmentRequest.Run();
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
