using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Text;
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

        public (Stream, AggregateException) Merge(string inputFile, Dictionary<string, string> fieldValues)
        {
            var exception = ValidateParameterInputFile(inputFile);
            if(exception!=null){return (Stream.Null, new AggregateException(exception));}
            //
            using (var inputStream = new FileInfo(inputFile).OpenRead())
            {
                var (result,exceptions) = MergeInternal(inputStream, fieldValues);
                return (result, new AggregateException(exceptions));
            }
        }

        public (Stream,AggregateException) Merge(Stream input, Dictionary<string,string> fieldValues)
        {
            var (result,exceptions) = MergeInternal(input, fieldValues);
            return (result, new AggregateException(exceptions));
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


        (Stream, List<Exception>) MergeInternal(Stream input, Dictionary<string, string> fieldValues)
        {
            fieldValues = LogAndEnsureFieldValues(fieldValues, new Dictionary<string, string>());
            var exceptions = ValidateParameters(input, fieldValues);
            if (exceptions.Any()){ return (Stream.Null, exceptions); }
            try
            {
                var outputMMF = MemoryMappedFile.CreateNew(null, Settings.MaximumOutputFileSize, MemoryMappedFileAccess.ReadWrite);
                var outputStream = outputMMF.CreateViewStream(0, input.Length);
                input.CopyTo(outputStream);
                return (outputStream, exceptions);
                //                using (var outputStream = new MemoryStream(Settings.MaximumMemoryStreamSize))
                //                {
                //                    input.CopyTo(outputStream);
                //                    return (outputStream, exceptions);
                //                }
            }
            catch (Exception e){ exceptions.Add(e); }
            return (Stream.Null, exceptions);
        }

        Dictionary<string, string> LogAndEnsureFieldValues(Dictionary<string, string> fieldValues,
                                                           Dictionary<string, string> @default)
        {
            if (fieldValues == null || fieldValues.Count == 0)
            {
                Logger.LogDebug("Starting Merge input stream with empty fieldValues={@fieldValues}", fieldValues);
            }
            else
            {
                Logger.LogTrace("Starting Merge input stream with fieldValues={@fieldValues}", fieldValues);
            }
            return fieldValues??@default;
        }


        List<Exception> ValidateParameters(Stream input, Dictionary<string, string> fieldValues)
        {
            var exceptions = new List<Exception>();
            if (input == null){exceptions.Add(new ArgumentNullException(nameof(input)));}
            if (fieldValues == null){exceptions.Add(new ArgumentNullException(nameof(fieldValues)));}
            return exceptions;
        }

        static Exception ValidateParameterInputFile(string inputFile)
        {
            if (inputFile == null) { return new ArgumentNullException(nameof(inputFile)); }
            if (!new FileInfo(inputFile).Exists) { return new FileNotFoundException("File Not Found: " + inputFile, inputFile); }

            return null;
        }
    }
}
