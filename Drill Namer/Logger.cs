using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Autodesk.AutoCAD.ApplicationServices;

namespace Drill_Namer
{
    public static class Logger
    {
        private static NLog.Logger logger;

        static Logger()
        {
            ConfigureLogger();
        }

        /// <summary>
        /// Configures the NLog logger with a dynamic log file path based on the current drawing.
        /// </summary>
        private static void ConfigureLogger()
        {
            try
            {
                // Create a new NLog configuration
                var config = new LoggingConfiguration();

                // Get the current drawing's directory for the log file path
                string logFilePath = GetLogFilePath();

                // Create a FileTarget with the dynamic log file path
                var logfile = new FileTarget("logfile")
                {
                    FileName = logFilePath,
                    Layout = "${longdate} | ${level:uppercase=true} | ${message}"
                };

                // Add the rule for mapping loggers to the FileTarget
                config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);

                // Apply the configuration
                LogManager.Configuration = config;

                logger = LogManager.GetCurrentClassLogger();
            }
            catch (Exception ex)
            {
                try
                {
                    string source = "DrillNamer";
                    string log = "Application";
                    if (!EventLog.SourceExists(source))
                    {
                        EventLog.CreateEventSource(source, log);
                    }
                    EventLog.WriteEntry(source, $"Error initializing NLog: {ex.Message}", EventLogEntryType.Error);
                }
                catch { }
            }
        }

        /// <summary>
        /// Retrieves the log file path based on the current drawing's directory.
        /// If the drawing is unsaved, defaults to the user's Documents folder.
        /// </summary>
        /// <returns>Full path to the log file.</returns>
        private static string GetLogFilePath()
        {
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                if (acDoc != null && acDoc.IsNamedDrawing && !string.IsNullOrEmpty(acDoc.Name))
                {
                    string drawingDirectory = Path.GetDirectoryName(acDoc.Name);
                    string logFilePath = Path.Combine(drawingDirectory, "DrillNamer.log");
                    return logFilePath;
                }
                else
                {
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "DrillNamer.log");
                }
            }
            catch
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "DrillNamer.log");
            }
        }

        /// <summary>
        /// Updates the log file path based on the current drawing's directory.
        /// This method is called before each logging action to ensure the log file is correctly located.
        /// </summary>
        private static void UpdateLogFilePath()
        {
            try
            {
                string logFilePath = GetLogFilePath();
                var config = LogManager.Configuration;
                var logfile = config.FindTargetByName<FileTarget>("logfile");
                if (logfile != null)
                {
                    // Use Render to obtain the actual file name
                    string currentFileName = logfile.FileName.Render(new LogEventInfo());
                    if (!string.Equals(currentFileName, logFilePath, StringComparison.OrdinalIgnoreCase))
                    {
                        logfile.FileName = logFilePath;
                        LogManager.ReconfigExistingLoggers();
                    }
                }
            }
            catch
            {
                // Silently fail
            }
        }

        public static void LogInfo(string message)
        {
            UpdateLogFilePath();
            logger.Info(message);
        }

        public static void LogWarning(string message)
        {
            UpdateLogFilePath();
            logger.Warn(message);
        }

        public static void LogError(string message)
        {
            UpdateLogFilePath();
            logger.Error(message);
        }

        public static void LogDebug(string message)
        {
            UpdateLogFilePath();
            logger.Debug(message);
        }
    }
}
