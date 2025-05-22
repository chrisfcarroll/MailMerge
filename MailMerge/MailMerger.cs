using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MailMerge.CommandLine;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using Settings = MailMerge.Properties.Settings;

namespace MailMerge
{
    /// <summary>
    /// A component for editing Word Docx files, and in particular for populating merge fields.
    /// Works with streams or files.
    ///
    /// <list type="bullet">
    ///     <item>
    ///         <term>Capabilities Implemented So Far</term>
    ///         <description><see cref="KnownWordProcessingMLTransformationsReadMe.AllKnown"/></description>
    ///     </item>
    ///     <item>
    ///         <term>Full Spec</term>
    ///         <description>https://www.ecma-international.org/publications/standards/Ecma-376.htm</description>
    ///     </item>
    ///     <item>
    ///         <term>Much more readable abbreviated spec</term>
    ///         <description>http://officeopenxml.com/</description>
    ///     </item>
    /// </list>
    /// </summary>
    public class MailMerger
    {
        public const string DATEKey = "DATE";
        internal readonly ILogger Logger;
        internal readonly Settings Settings;
        internal readonly bool MatchFieldNamesCaseSensitively;

        /// <summary>
        /// Use this property for DATE substitions if you require rudimentary WordProcessingML MergeFormat handling.
        /// Otherwise, a simpler choice is to add an entry with <code>Key=</code><seealso cref="DATEKey"/> to the fieldValues 
        /// passed in to <seealso cref="Merge(string,Dictionary{string,string})"/>
        /// </summary>
        public DateTime? DateTime { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="logger">a logger. If you pass null, then <see cref="NullLogger.Instance"/> will be used.</param>
        /// <param name="settings"></param>
        /// <param name="dateTime"></param>
        /// <param name="matchFieldNamesCaseSensitively"></param>
        public MailMerger(ILogger logger, Settings settings=null, DateTime? dateTime = null, bool matchFieldNamesCaseSensitively=false)
        {
            Logger = logger?? NullLogger.Instance; 
            Settings = settings??new Settings();
            MatchFieldNamesCaseSensitively = matchFieldNamesCaseSensitively;
            DateTime = dateTime;
        }
        
        /// <summary>Create a new MailMerger with Logger and Settings from <see cref="Startup"/></summary>
        public MailMerger()
        {
            Startup.Configure();
            Logger = Startup.CreateLogger<MailMerger>();
            Settings = Startup.Settings;
            DateTime=null;
        }

        /// <summary>
        /// Open the given <paramref name="inputDocxFileName"/> and merge fieldValues into it.
        /// </summary>
        /// <returns>
        /// Item1: The merged result as a new stream. If <paramref name="fieldValues"/> is empty, a <em>copy</em> of the original document is returned.
        /// Item2: An <see cref="AggregateException"/> containing any exceptions that were raised during the process.
        /// </returns>
        /// <param name="inputDocxFileName">Input file.</param>
        /// <param name="fieldValues">A dictionary keyed on mergefield names used in the document.</param>
        /// <remarks>Error handling: 
        /// You can inspect the returned <seealso cref="AggregateException.InnerExceptions"/> , or simply <code>throw</code it.</remarks>
        public (Stream, AggregateException) Merge(string inputDocxFileName, Dictionary<string, string> fieldValues)
        {
            var exception = ValidateParameterInputFile(inputDocxFileName);
            if(exception!=null){return (Stream.Null, new AggregateException(exception));}
            //
            using (var inputStream = new FileInfo(inputDocxFileName).OpenRead())
            {
                var (result,exceptions) = MergeInternal(inputStream, fieldValues);
                return (result, new AggregateException(exceptions));
            }
        }
        /// <summary>
        /// Open the given <paramref name="inputDocxFileName"/> and merge fieldValues into it. Save the result to <paramref name="outputDocxFileName"/>
        /// </summary>
        /// <returns>
        /// Item1: true if we saved an output document, false otherwise.
        /// Item2: An <see cref="AggregateException"/> containing any exceptions that were raised during the process.
        /// </returns>
        /// <param name="inputDocxFileName">Input file.</param>
        /// <param name="fieldValues">A dictionary keyed on mergefield names used in the document.</param>
        public (bool, AggregateException) Merge(string inputDocxFileName, Dictionary<string, string> fieldValues, string outputDocxFileName)
        {
            var exceptioni = ValidateParameterInputFile(inputDocxFileName);
            var exceptiono = ValidateParameterOutputFile(outputDocxFileName);
            if (exceptioni != null) { return (false, new AggregateException( new[] { exceptioni, exceptiono }.Where(e=>e!=null) )); }
            //
            try
            {
                using (var outstream = new FileInfo(outputDocxFileName).OpenWrite())
                using (var inputStream = new FileInfo(inputDocxFileName).Open(FileMode.Open, FileAccess.Read, FileShare.Read) )
                {
                    var (result, exceptions) = MergeInternal(inputStream, fieldValues);
                    result.Position = 0;
                    result.CopyTo(outstream);
                    return (true, new AggregateException(exceptions));
                }
            }
            catch(Exception e){ return (false, new AggregateException(e)); }
        }

        /// <summary>
        /// Open the given <paramref name="input"/> stream as a docx file, and merge fieldValues into it.
        /// </summary>
        /// <returns>
        /// Item1: The merged result as a new stream. If <paramref name="fieldValues"/> is empty, a <em>copy</em> of the original stream is returned.
        /// Item2: An <see cref="AggregateException"/> containing any exceptions that were raised during the process.
        /// </returns>
        /// <param name="input">Input stream of a docx file.</param>
        /// <param name="fieldValues">A dictionary keyed on mergefield names used in the document.</param>
        public (Stream,AggregateException) Merge(Stream input, Dictionary<string,string> fieldValues)
        {
            var (result,exceptions) = MergeInternal(input, fieldValues);
            return (result, new AggregateException(exceptions));
        }

        /// <summary>
        /// Open the given <paramref name="input"/> stream as a docx file, merges fieldValues into it, and saves the result to <paramref name="outputPath"/>
        /// </summary>
        /// <returns>
        /// Item1: The merged result as a new stream. If <paramref name="fieldValues"/> is empty, a <em>copy</em> of the original stream is returned.
        /// Item2: An <see cref="AggregateException"/> containing any exceptions that were raised during the process.
        /// </returns>
        /// <param name="input">Input stream of a docx file.</param>
        /// <param name="fieldValues">A dictionary keyed on mergefield names used in the document.</param>
        public (bool, AggregateException) Merge(Stream input, Dictionary<string,string> fieldValues, string outputPath)
        {
            var (result,exceptions) = MergeInternal(input, fieldValues);
            if (result != null) try
                {
                    using (var outstream = new FileInfo(outputPath).Create())
                    {
                        result.Position = 0;
                        result.CopyToAsync(outstream);
                        return (true, new AggregateException(exceptions));
                    }
                }
                catch (Exception e){ exceptions.Add(e); }
            return (false, new AggregateException(exceptions));
        }


        /// <summary>First copies the input stream to the output stream, then edits the outputstream to apply any mergefield transforms</summary>
        /// <returns>The edited output stream</returns>
        /// <param name="input">a docx file </param>
        /// <param name="fieldValues">A dictionary keyed on mergefield names used in the document.</param>
        (Stream, List<Exception>) MergeInternal(Stream input, Dictionary<string, string> fieldValues)
        {
            var exceptions = ValidateParameters(input, fieldValues);
            if (exceptions.Any() ) { return (Stream.Null, exceptions); }
            try
            {
                var outputStream= new MemoryStream(); 
                input.CopyTo(outputStream);                
                fieldValues = LogAndEnsureFieldValues(fieldValues, new Dictionary<string, string>());

                if (fieldValues.Any())
                {
                    if (!MatchFieldNamesCaseSensitively && !new[] {StringComparer.InvariantCultureIgnoreCase, StringComparer.CurrentCultureIgnoreCase, StringComparer.OrdinalIgnoreCase}.Contains(fieldValues.Comparer))
                    {
                        fieldValues=new Dictionary<string,string>(fieldValues,StringComparer.CurrentCultureIgnoreCase);
                    }
                    ApplyAllKnownMergeTransformationsToMainDocumentPart(fieldValues, outputStream);
                }
                
                return (outputStream, exceptions);
            }
            catch (Exception e){ exceptions.Add(e); }
            return (Stream.Null, exceptions);
        }

        /// <summary>Use the power of OOXml to edit the given <paramref name="editableWPXmlStream"/>
        /// by applying all known merge transformations. As at June 2023, all known transformations
        /// are <see cref="KnownWordProcessingMLTransformationsReadMe.AllKnown"/>
        /// </summary>
        /// <param name="fieldValues">The merge field replacements dictionary</param>
        /// <param name="editableWPXmlStream">an editable stream containing a WordProcessingML document</param>
        internal void ApplyAllKnownMergeTransformationsToMainDocumentPart(Dictionary<string, string> fieldValues, Stream editableWPXmlStream)
        {
            var xdoc = GetMainDocumentPartXml(editableWPXmlStream);

            xdoc.MergeSimpleMergeFields(fieldValues, Logger);
            xdoc.MergeComplexMergeFields(fieldValues,Logger);
            var dateValueForMerge = fieldValues.ContainsKey(DATEKey) ? fieldValues[DATEKey] : DateTime?.ToLongDateString();
            xdoc.MergeDateFields(Logger, DateTime, dateValueForMerge);

            using (var wpDocx = WordprocessingDocument.Open(editableWPXmlStream, true))
            {
                var bodyNode = xdoc.SelectSingleNode("/w:document/w:body", OoXmlNamespace.Manager);
                var documentBody = new Body(bodyNode.OuterXml);
                wpDocx.MainDocumentPart.Document.Body = documentBody;

                if (wpDocx.MainDocumentPart?.HeaderParts is not null)
                {
                    foreach (var headerPart in wpDocx.MainDocumentPart.HeaderParts)
                    {
                        var headerXDoc = new XmlDocument(OoXmlNamespace.Manager.NameTable);

                        using (var partStream = headerPart.GetStream(FileMode.Open, FileAccess.Read))
                        {
                            headerXDoc.Load(partStream);
                        }

                        // Apply merge transformations
                        headerXDoc.MergeSimpleMergeFields(fieldValues, Logger);
                        headerXDoc.MergeComplexMergeFields(fieldValues, Logger);
                        headerXDoc.MergeDateFields(Logger, DateTime, dateValueForMerge);

                        // Save modified XML back to the header part
                        using (var partStream = headerPart.GetStream(FileMode.Create, FileAccess.Write))
                        {
                            headerXDoc.Save(partStream);
                        }
                    }
                }
            }
        }

        public static XmlDocument GetMainDocumentPartXml(Stream docxStream)
        {
            var xdoc = new XmlDocument(OoXmlNamespace.Manager.NameTable);
            using (var wpDocx = WordprocessingDocument.Open(docxStream, false))
            using (var docOutStream = wpDocx.MainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
            {
                xdoc.Load(docOutStream);
            }
            return xdoc;
        }

        Dictionary<string, string> LogAndEnsureFieldValues(Dictionary<string, string> fieldValues, Dictionary<string, string> @default)
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
            try
            {
                if (inputFile == null) { return new ArgumentNullException(nameof(inputFile)); }
                if (!new FileInfo(inputFile).Exists) { return new FileNotFoundException("File Not Found: " + inputFile, inputFile); }
            }
            catch (Exception e) { return e; }

            return null;
        }

        static Exception ValidateParameterOutputFile(string outputFile)
        {
            try
            {
                if (outputFile == null) { return new ArgumentNullException(nameof(outputFile)); }
                var outputfileinfo = new FileInfo(outputFile);
                if (outputfileinfo.Exists) { return new IOException("File already exists : " + outputFile); }
                try { using (var test = outputfileinfo.OpenWrite()) { test.Close(); } } finally { outputfileinfo.Delete(); }
            }
            catch (Exception e) { return e; }

            return null;
        }
    }
}
