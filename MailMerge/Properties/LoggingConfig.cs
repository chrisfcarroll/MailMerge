using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Events;

namespace MailMerge.Properties
{
    static class LoggingConfig
    {
        static string LoggingConfigurationSectionName = "Logging";
        static string LogLevelSettingName = "LogLevel";
        
        public static ILoggerFactory FromConfiguration(this ILoggerFactory loggerFactory, IConfiguration configuration)
        {
            var logconfig = new LoggerConfiguration();
            var logLevelSetting = configuration.GetSection(LoggingConfigurationSectionName)[LogLevelSettingName];
            
            if (!Enum.TryParse<LogEventLevel>(logLevelSetting, out var serilogLevel))
            {
                if (Enum.TryParse<LogLevel>(logLevelSetting, out var mslogLevel))
                {
                    serilogLevel = MsLevelToSerilogLevel[mslogLevel];
                    if (mslogLevel == LogLevel.None) { logconfig = logconfig.Filter.ByExcluding(x => true); }
                }
                else
                {
                    serilogLevel = LogEventLevel.Information;
                }
            }

            return loggerFactory.AddSerilog(logconfig.MinimumLevel.Is(serilogLevel).WriteTo.Console().CreateLogger());
        }

        static readonly Dictionary<LogLevel,LogEventLevel> MsLevelToSerilogLevel= new Dictionary<LogLevel,LogEventLevel>
        {
            {LogLevel.Critical,LogEventLevel.Fatal},
            {LogLevel.Debug,LogEventLevel.Debug},
            {LogLevel.Error,LogEventLevel.Error},
            {LogLevel.Information,LogEventLevel.Information},
            {LogLevel.None,LogEventLevel.Fatal},
            {LogLevel.Trace,LogEventLevel.Verbose},
            {LogLevel.Warning,LogEventLevel.Warning},
        };
    }
}