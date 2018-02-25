using System.Collections.Generic;
using System.IO;
using MailMerge.OoXml.Properties;
using Microsoft.Extensions.Logging;

namespace MailMerge.OoXml
{
    public class MailMerge
    {
        internal readonly ILogger Logger;
        internal readonly Settings Settings;

        public MailMerge(ILogger logger, Settings settings)
        {
            Logger = logger;
            Settings = settings;
        }

        public Stream Merge(Stream input, Dictionary<string,string> fieldValues)
        {
            Logger.LogInformation("This is a console-runnable component with SomeSetting={@SomeSetting}", Settings.SomeSetting);

            return input;
        }
        
        public bool Merge(Stream input, Dictionary<string,string> fieldValues, FileInfo outputPath)
        {
            Logger.LogInformation("This is a console-runnable component with SomeSetting={@SomeSetting}", Settings.SomeSetting);

            return true;
        }
    }
}
