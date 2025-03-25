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

        private static void SetColor(ConsoleColor color)
        {
            if (Console.ForegroundColor != color)
            {
                Console.ForegroundColor = color;
            }
        }

        private static void ResetColor()
        {
            if (Console.ForegroundColor != ConsoleColor.Gray)
            {
                Console.ResetColor();
            }
        }

        public static void Write(string content, ConsoleColor color)
        {
            EnsurePosition();
            SetColor(color);
            Console.Write(content);
            ResetColor();
            lastLine = content;
        }

        public static void WriteLine(string content, ConsoleColor color)
        {
            EnsurePosition();
            SetColor(color);
            Console.WriteLine(content);
            ResetColor();
            lastLine = content;
        }

        public static bool ChoiceYesNo(string content, ConsoleColor color)
        {
            SetColor(color);

            Console.Write(content + " (Y/N) : ");
            char response = Char.ToUpper(Console.ReadKey().KeyChar);

            Console.Write('\n');
            ResetColor();
            return response == 'Y';
        }

        public static int PromptInt(string content, ConsoleColor color)
        {
            SetColor(color);

            Console.Write(content + ": ");
            string? input = Console.ReadLine();

            Console.Write('\n');
            ResetColor();
            return int.TryParse(input, out int value) ? value : -1;
        }

        public static string PromptForPath(ConsoleColor color, string content = "Please enter in a path")
        {
            SetColor(color);

            Console.Write(content + ": ");
            string? input = Console.ReadLine();

            // Remove quotes
            input = input?.Trim('"');

            Console.Write('\n');
            ResetColor();
            return input ?? string.Empty;
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
            SetColor(barColor);

            // Set cursor down when we reach 100%.
            if (progressPercent >= progressMax)
            {
                Console.Write('\n');
                ResetColor();
            }

            // Compare last line to the new constructed line
            if (progressBar == lastLine)
            {
                return;
            }
            lastLine = progressBar;

            // Write the progress bar ensuring each character has the correct color.
            Write(progressBar, barColor);

            Sw.Restart(); // this can also start the stopwatch
        }

        public static void ClearLine()
        {
            if (lastLine != null)
            {
                Console.CursorLeft = 0;
                Console.Write(new string(' ', lastLine.Length));
                Console.CursorLeft = 0;
                lastLine = string.Empty;
            }
        }
    }
}
