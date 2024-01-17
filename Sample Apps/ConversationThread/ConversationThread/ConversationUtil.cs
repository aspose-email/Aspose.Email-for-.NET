using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace ConversationThread
{
    /// <summary>
    /// Implements the Aspose.Email's PST/MAPI functionality to search conversation related messages.
    /// The search is based on a PidTagConversationIndex property header comparison. 
    /// Related messages have the same header in this property.
    /// Visit https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/9e994fbb-b839-495f-84e3-2c8c02c7dd9b 
    /// for a more detailed property description. 
    /// You can complicate the implementation by adding the conversation messages hierarchical sorting for example.
    /// </summary>
    public class ConversationUtil : IDisposable
    {
        private PersonalStorage pst;
        private readonly FolderInfo folder;
        private readonly Action action;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConversationUtil"/> class.
        /// Open the pst file and get the folder in which we'll be searching.
        /// </summary>
        /// <param name="pstFileName">The PST file name.</param>
        /// <param name="standardFolder">The standard folder (Inbox, Sent Items etc.)</param>
        /// <param name="callback">Callback function required to indicate the process.</param>
        public ConversationUtil(string pstFileName, StandardIpmFolder standardFolder, Action callback)
        {
            pst = PersonalStorage.FromFile(pstFileName);
            folder = pst.GetPredefinedFolder(standardFolder);
            action = callback;
        }

        /// <summary>
        /// Traverses the specified folder and returns list of coversation related message IDs.
        /// </summary>
        /// <returns>The <see cref="MessageThread"/> collection</returns>
        public List<MessageThread> GroupMessagesByThread()
        {
            // Result
            var list = new List<MessageThread>();

            foreach (var entryId in folder.EnumerateMessagesEntryId())
            {
                // Extract PidTagConversationIndex property and its header value (22 bytes).
                var convIndexProperty = pst.ExtractProperty(Convert.FromBase64String(entryId), KnownPropertyList.ConversationIndex.Tag);
                var convIndexHeader = (convIndexProperty != null && convIndexProperty.Data != null) ? convIndexProperty.Data.Take(22).ToArray() : null;
                var convIndexHeaderString = convIndexHeader != null ? Convert.ToBase64String(convIndexHeader) : null;

                if (convIndexHeaderString != null)
                {
                    var messageThread = list.Find(x => x.ConversationIndex == convIndexHeaderString);

                    if (messageThread == null)
                    {
                        messageThread = new MessageThread(convIndexHeaderString);
                        messageThread.Messages.Add(entryId);
                        list.Add(messageThread);
                    }
                    else
                    {
                        messageThread.Messages.Add(entryId);
                    }
                }

                action();
            }

            return list.Where(x => x.Messages.Count > 1).ToList();
        }

        /// <summary>
        /// Extracts and saves converation messages.
        /// Related messages will be in the same folder.
        /// </summary>
        /// <param name="list">The collection of <see cref="MessageThread"/>.</param>
        /// <param name="outDirectory">The output path to save.</param>
        public void SaveThreads(List<MessageThread> list, string outDirectory)
        {
            var invalidChars = Path.GetInvalidFileNameChars();

            foreach (var item in list)
            {
                string? threadDirectory = null;

                var i = item.Messages.Count - 1;

                foreach (var entryId in item.Messages)
                {
                    var msg = pst.ExtractMessage(entryId);

                    if (threadDirectory == null)
                    {
                        var threadName = new string(msg.ConversationTopic.Where(x => !invalidChars.Contains(x)).ToArray()).Trim();
                        threadDirectory = Path.Combine(outDirectory, threadName);
                        Directory.CreateDirectory(threadDirectory);
                    }

                    msg.Save(Path.Combine(threadDirectory, $"{i--}.msg"));
                }

                threadDirectory = null;

                action();
            }
        }

        public void Dispose()
        {
            if (pst != null)
            {
                pst.Dispose();
                pst = null;
            }
        }
    }
}
