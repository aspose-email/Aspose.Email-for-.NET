
namespace ConversationThread
{
    /// <summary>
    /// Represents messages within a conversation thread.
    /// </summary>
    public class MessageThread
    {
        private string convIndexHeader;
        private List<string> messages;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageThread"/> class.
        /// </summary>
        /// <param name="convIndexHdr">The Base64 string representation of conversation cndex header (22 bytes).</param>
        public MessageThread(string convIndexHdr)
        {
            this.convIndexHeader = convIndexHdr;
            this.messages = new List<string>();
        }

        /// <summary>
        /// Gets or sets the Base64 string representation of conversation cndex header (22 bytes).
        /// </summary>
        public string ConversationIndex { get => convIndexHeader; set => convIndexHeader = value; }

        /// <summary>
        /// Gets a list of message IDs that are within a conversation thread.
        /// </summary>
        public List<string> Messages { get => messages; }
    }
}
