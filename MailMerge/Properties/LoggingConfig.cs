using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace MailMerge.Properties
{
    static class LoggingConfig
    {
        public static Instance FromConfig;
        const string LoggingConfigurationSectionName = "Logging";

        public static ILoggerFactory FromConfiguration(this ILoggerFactory loggerFactory, IConfiguration configuration, ILoggerProvider provider=null)
        {
            var configurationSection = configuration.GetSection(LoggingConfigurationSectionName);
            configurationSection.Bind(FromConfig=new Instance());
            loggerFactory = loggerFactory ?? new FallbackLoggerFactory();
            loggerFactory.AddProvider( FromConfig.Provider= provider??new FallbackLoggerProvider());
            return loggerFactory;
        }

        public class Instance
        {
            public string LogLevel { get; set; } = "Information";
            public ILoggerProvider Provider { get; set; }
        }
    }
}