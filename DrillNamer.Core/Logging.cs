using NLog;
using NLog.Config;
using NLog.Targets;

namespace DrillNamer.Core;

/// <summary>
/// Provides a configured NLog logger.
/// </summary>
public static class Logging
{
    private static readonly Logger Logger;

    static Logging()
    {
        var config = new LoggingConfiguration();
        var fileTarget = new FileTarget("logfile")
        {
            FileName = "${specialfolder:folder=MyDocuments}/DrillNamer.log",
            Layout = "${longdate} | ${level:uppercase=true} | ${message}"
        };
        config.AddRule(LogLevel.Info, LogLevel.Fatal, fileTarget);
        LogManager.Configuration = config;
        Logger = LogManager.GetLogger("DrillNamer");
    }

    public static void Info(string message) => Logger.Info(message);
    public static void Warn(string message) => Logger.Warn(message);
    public static void Error(string message) => Logger.Error(message);
    public static void Debug(string message) => Logger.Debug(message);
}
