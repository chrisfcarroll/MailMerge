using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

        public (Stream,AggregateException) Merge(Stream input, Dictionary<string,string> fieldValues)
        {
            var (result,exceptions) = MergeInternal(input, fieldValues);
            return (result, new AggregateException(exceptions));
        }

        (Stream,List<Exception>) MergeInternal(Stream input, Dictionary<string, string> fieldValues)
        {
            var exceptions= new List<Exception>();
            try
            {
                Logger.LogTrace("Starting Merge input stream with fieldValues={@fieldValues}", fieldValues);
                exceptions = ValidateParameters(input, fieldValues);
                if (exceptions.Any()){ return (Stream.Null, exceptions); }
                //

                //
            }
            catch (Exception e){ exceptions.Add(e); }
            return (Stream.Null, exceptions);
        }

        public (bool, AggregateException) Merge(Stream input, Dictionary<string,string> fieldValues, string outputPath)
        {
            var (result,exceptions) = MergeInternal(input, fieldValues);
            if (result != null) try
            {
                using (var outstream = new FileInfo(outputPath).Create())
                {
                    result.CopyToAsync(outstream);
                    return (true, new AggregateException(exceptions));
                }
            }
            catch (Exception e){ exceptions.Add(e); }
            return (false, new AggregateException(exceptions));
        }

        List<Exception> ValidateParameters(Stream input, Dictionary<string, string> fieldValues)
        {
            var exceptions = new List<Exception>();
            if (input == null){exceptions.Add(new ArgumentNullException(nameof(input)));}
            if (fieldValues == null){exceptions.Add(new ArgumentNullException(nameof(fieldValues)));}
            return exceptions;
        }
    }
}
