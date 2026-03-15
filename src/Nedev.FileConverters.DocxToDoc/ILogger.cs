using System;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Defines a logger for recording conversion operations and errors.
    /// </summary>
    public interface ILogger
    {
        /// <summary>
        /// Logs a debug message.
        /// </summary>
        void LogDebug(string message);

        /// <summary>
        /// Logs an informational message.
        /// </summary>
        void LogInfo(string message);

        /// <summary>
        /// Logs a warning message.
        /// </summary>
        void LogWarning(string message);

        /// <summary>
        /// Logs an error message.
        /// </summary>
        void LogError(string message);

        /// <summary>
        /// Logs an error message with exception details.
        /// </summary>
        void LogError(string message, Exception exception);
    }

    /// <summary>
    /// A no-operation logger that discards all log messages.
    /// </summary>
    public class NullLogger : ILogger
    {
        public static readonly NullLogger Instance = new NullLogger();

        public void LogDebug(string message) { }
        public void LogInfo(string message) { }
        public void LogWarning(string message) { }
        public void LogError(string message) { }
        public void LogError(string message, Exception exception) { }
    }

    /// <summary>
    /// A simple console logger for debugging purposes.
    /// </summary>
    public class ConsoleLogger : ILogger
    {
        public void LogDebug(string message)
        {
            Console.WriteLine($"[DEBUG] {message}");
        }

        public void LogInfo(string message)
        {
            Console.WriteLine($"[INFO] {message}");
        }

        public void LogWarning(string message)
        {
            Console.WriteLine($"[WARN] {message}");
        }

        public void LogError(string message)
        {
            Console.WriteLine($"[ERROR] {message}");
        }

        public void LogError(string message, Exception exception)
        {
            Console.WriteLine($"[ERROR] {message}");
            Console.WriteLine($"Exception: {exception.Message}");
            Console.WriteLine($"Stack Trace: {exception.StackTrace}");
        }
    }
}
