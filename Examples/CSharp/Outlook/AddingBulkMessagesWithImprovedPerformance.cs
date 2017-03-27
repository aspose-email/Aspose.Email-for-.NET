using System.IO;
using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

/* This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Email for .NET 
   API reference when the project is build. Please check https://Docs.nuget.org/consume/nuget-faq 
   for more information. If you do not wish to use NuGet, you can manually download 
   Aspose.Email for .NET API from http://www.aspose.com/downloads, 
   install it and then add its reference to this project. For any issues, questions or suggestions 
   please feel free to contact us using http://www.aspose.com/community/forums/default.aspx            
*/

namespace Aspose.Email.Examples.CSharp.Email.Outlook
{
    class AddingBulkMessagesWithImprovedPerformance
    {
        public static void Run()
        {
            // ExStart:AddingBulkMessagesWithImprovedPerformance
            // The path to the file directory.
            string dataDir = RunExamples.GetDataDir_Outlook();

            // Load the Outlook file
            string path = dataDir + "PersonalStorageFile2.pst";
            AddMessagesInBulkMode(path, "Contacts");
        }

        // ExStart:AddingBulkMessages
        private static void AddMessagesInBulkMode(string fileName, string msgFolderName)
        {
            using (PersonalStorage personalStorage = PersonalStorage.FromFile(fileName))
            {
                FolderInfo folder = personalStorage.RootFolder.GetSubFolder("myInbox");
                folder.MessageAdded += OnMessageAdded;
                folder.AddMessages(new MapiMessageCollection(msgFolderName));
            }
        }
        static void OnMessageAdded(object sender, MessageAddedEventArgs e)
        {
            Console.WriteLine(e.EntryId);
            Console.WriteLine(e.Message.Subject);
        }
        // ExEnd:AddingBulkMessages

        // ExStart:IEnumerableImplementation
        public class MapiMessageCollection : IEnumerable<MapiMessage>
        {
            private string path;

            public MapiMessageCollection(string path)
            {
                this.path = path;
            }

            public IEnumerator<MapiMessage> GetEnumerator()
            {
                return new MapiMessageEnumerator(path);
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }
        }

        public class MapiMessageEnumerator : IEnumerator<MapiMessage>
        {
            private readonly string[] files;

            private int position = -1;

            public MapiMessageEnumerator(string path)
            {
                string path1 = RunExamples.GetDataDir_Outlook();
                files = Directory.GetFiles(path1);
            }

            public bool MoveNext()
            {
                position++;
                return (position < files.Length);
            }

            public void Reset()
            {
                position = -1;
            }

            object IEnumerator.Current
            {
                get
                {
                    return Current;
                }
            }

            public MapiMessage Current
            {
                get
                {
                    try
                    {
                        return MapiMessage.FromFile(files[position]);
                    }
                    catch (IndexOutOfRangeException)
                    {
                        throw new InvalidOperationException();
                    }
                }
            }
            public void Dispose()
            {
            }
        }
        // ExEnd:IEnumerableImplementation
    }
}
