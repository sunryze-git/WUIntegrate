using System.Diagnostics;

namespace WUIntegrate
{
    static class ConsoleWriter
    {
        private static string lastLine = string.Empty;
        private static readonly Stopwatch Sw = new();

        private static void EnsurePosition()
        {
            // ensure the console cursor is on a blank line and all the way to the left
            if (Console.CursorLeft != 0)
            {
                Console.CursorLeft = 0;
            }
        }
        public static void Write(string content, ConsoleColor color)
        {
            EnsurePosition();
            Console.ForegroundColor = color;
            Console.Write(content);
            Console.ResetColor();
            lastLine = content;
        }

        public static void WriteLine(string content, ConsoleColor color)
        {
            EnsurePosition();
            Console.ForegroundColor = color;
            Console.WriteLine(content);
            Console.ResetColor();
            lastLine = content;
        }

        public static bool ChoiceYesNo(string content, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.Write(content + " (Y/N) : ");
            char response = Char.ToUpper(Console.ReadKey().KeyChar);
            Console.Write('\n');
            Console.ResetColor();
            return response == 'Y';
        }

        public static int PromptInt(string content, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.Write(content + ": ");
            string? input = Console.ReadLine();
            Console.Write('\n');
            Console.ResetColor();
            return int.Parse(input ?? "0");
        }

        public static void WriteProgress(int maxChars, int progress, int progressMax, string message, ConsoleColor barColor)
        {
            var progressPercent = (int)((double)progress / progressMax * 100);

            if (Sw.IsRunning && progressPercent < 100 && Sw.ElapsedMilliseconds < 100)
            {
                return;
            }

            var emptyChars = new string('-', maxChars - progressPercent);
            var filledChars = new string('#', progressPercent);

            var progressBar = $"|{filledChars}{emptyChars}| {progressPercent}% {message}";

            // Save time
            if (Console.ForegroundColor != barColor)
            {
                Console.ForegroundColor = barColor;
            }

            // Set cursor down when we reach 100%.
            if (progressPercent >= progressMax)
            {
                Console.CursorTop++;
                Console.CursorLeft = 0;
            }

            // Compare last line to the new constructed line
            if (progressBar == lastLine)
            {
                return;
            }
            lastLine = progressBar;

            // Write the progress bar ensuring each character has the correct color.
            Console.CursorLeft = 0;
            Write(progressBar, barColor);

            if (!Sw.IsRunning)
            {
                Sw.Start();
            }
            else
            {
                Sw.Restart();
            }
        }

        public static void ClearLine()
        {
            Console.CursorLeft = 0;
            Console.Write(new string(' ', Console.WindowWidth - 1));
            Console.CursorLeft = 0;
        }
    }
}
