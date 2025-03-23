using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WUIntegrate
{
    public class Logger()
    {
        public static bool EnableLogging { get; set; }
        public static string? LogFile { get; set; }
        private static string LogTime
        {
            get
            {
                var time = DateTime.Now;
                return time.ToString("yyyy-MM-dd HH:mm:ss");
            }
        }

        /// <summary>
        /// Logs a message to the log file without displaying to console.
        /// </summary>
        /// <param name="message"></param>
        public static void Log(string message)
        {
            var logMsg = $"{LogTime} | [Info] | {message}";
            WriteToLog(logMsg);
        }

        /// <summary>
        /// Logs a message to the console and log file.
        /// </summary>
        /// <param name="message"></param>
        public static void Msg(string message)
        {
            var logMsg = $"{LogTime} | [Info] | {message}";
            ConsoleWriter.WriteLine(logMsg, ConsoleColor.Cyan);
            WriteToLog(logMsg);
        }

        /// <summary>
        /// Logs a warning message to the console and log file.
        /// </summary>
        /// <param name="message"></param>
        public static void Warn(string message)
        {
            var logMsg = $"{LogTime} | [Info] | {message}";
            ConsoleWriter.WriteLine(logMsg, ConsoleColor.Yellow);
            WriteToLog(logMsg);
        }

        /// <summary>
        /// Logs an error message to the console and log file.
        /// </summary>
        /// <param name="message"></param>
        public static void Error(string message)
        {
            var logMsg = $"{LogTime} | [Info] | {message}";
            ConsoleWriter.WriteLine(logMsg, ConsoleColor.Red);
            WriteToLog(logMsg);
        }

        /// <summary>
        /// Logs a fatal crash message to the console and log file.
        /// </summary>
        /// <param name="ex"></param>
        public static void Crash(Exception ex)
        {
            var logMsg = 
                $"""
                FATAL CRASH @ {LogTime}
                    ERROR: {ex.Message}
                    SOURCE: {ex.Source}
                    STACKTRACE: {ex.StackTrace}
                """;
            ConsoleWriter.WriteLine(logMsg, ConsoleColor.Magenta);
            WriteToLog(logMsg);
        }

        private static void WriteToLog(string line)
        {
            if (!EnableLogging)
            {
                return;
            }

            if (LogFile is null)
            {
                ConsoleWriter.WriteLine($"[E] Cannot write to log, log file path is null", ConsoleColor.Red);
                return;
            }

            // Write string to newline in log file
            try
            {
                using StreamWriter writer = new(LogFile, true);
                writer.WriteLine(line);
                writer.Flush();
            }
            catch (PathTooLongException)
            {
                ConsoleWriter.WriteLine($"[E] Cannot write to log, log file path is too long", ConsoleColor.Red);
            }
            catch (DirectoryNotFoundException)
            {
                ConsoleWriter.WriteLine($"[E] Cannot write to log, directory was not found.", ConsoleColor.Red);
            }
            catch (IOException)
            {
                ConsoleWriter.WriteLine($"[E] Cannot write to log, there was an IO exception.", ConsoleColor.Red);
            }
            catch (UnauthorizedAccessException)
            {
                ConsoleWriter.WriteLine($"[E] Cannot write to log, access was denied.", ConsoleColor.Red);
            }
        }
    }
}
