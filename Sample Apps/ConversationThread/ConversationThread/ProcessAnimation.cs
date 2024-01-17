namespace ConversationThread
{
    // A simple progress indicator.
    // Just for auxiliary uses.
    internal class ProcessAnimation
    {
        private string[] _animation = { "▄▀", "▄▄", "▀▄", "▄▀", "▀▀", "▀▄" };
        private (int Left, int Top) curPos;
        private int progress;

        private ProcessAnimation()
        {
            curPos = Console.GetCursorPosition();
            Console.CursorVisible = false;
            progress = 0;
        }

        public static ProcessAnimation Create()
        {
            return new ProcessAnimation();
        }

        public void Write()
        {
            Console.SetCursorPosition(curPos.Left + 2, curPos.Top);
            Console.Write(_animation[progress % _animation.Length]);
            progress++;
        }

        public void Success()
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.SetCursorPosition(curPos.Left + 2, curPos.Top);
            Console.Write("Done!");
            Console.ResetColor();
        }

        public void Fail()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.SetCursorPosition(curPos.Left + 2, curPos.Top);
            Console.Write("Error!");
            Console.ResetColor();
        }
    }
}
