using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using MailMerge.CommandLine;
using NUnit.Framework;
using TestBase;

namespace MailMerge.Tests.FunctionalSpecs
{
    [TestFixture]
    public class GivenFldSimpleMergeField
    {
        const string TestDocDir = "TestDocuments";
        const string DocWithSimpleMergeFieldDocx = "DocWithSimpleMergeField.docx";
        const string DocWithFieldInHeaderDocx = "DocWithFieldInHeader.docx";
        MailMerger sut;

        static Dictionary<string, string> SingleLine = new Dictionary<string, string>
        {
            {"SimpleMergeField","SimpleMergeField Was Replaced with a Single Line"},
        };

        static Dictionary<string, string> MultiLine = new Dictionary<string, string>
        {
            {"SimpleMergeField","SimpleMergeField Was\nReplaced\nwith multiple lines\ndelimited by newlines/"}
        };

        static readonly Dictionary<string, string> HeaderAndBodyFields = new Dictionary<string, string>
        {
            { "FieldInHeader", "Yes, you can!" }, // Expected in header
            { "FirstName", "A replaced first name field" }, // Expected in body
            { "LastName", "A replaced last name field" } // Expected in body
        };

        [TestCase(DocWithSimpleMergeFieldDocx, nameof(SingleLine))]
        [TestCase(DocWithSimpleMergeFieldDocx, nameof(MultiLine))]
        [TestCase(DocWithFieldInHeaderDocx, nameof(HeaderAndBodyFields))]
        public void Returns_TheDocumentWithMergeFieldsReplaced(string source, string sourceFieldsSource)
        {
            source = Path.Combine(TestDocDir, source);
            var sourceFields = GetType()
                              .GetField(sourceFieldsSource,BindingFlags.Static|BindingFlags.NonPublic)
                              .GetValue(this) as Dictionary<string,string>;

            Stream output = null; AggregateException exceptions;

            using (var original = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            try
            {
                (output, exceptions) = sut.Merge(original, sourceFields);

                var (outputText, outputXml) = output.AsWordDocumentMainPartTextAndXml();

                var (headerPartText, headerXml) = output.GetCombinedHeadersTextAndXml();
                
                var combinedDocumentText = outputText + " " + headerPartText;
                using (var outFile = new FileInfo(source.Replace(".", $" {sourceFieldsSource} Output.")).OpenWrite())
                {
                    output.Position = 0;
                    output.CopyTo(outFile);
                }

                if (exceptions.InnerExceptions.Any()) { throw exceptions; }

                var combinedOutputXml = outputXml + " " + headerXml;

                sourceFields
                    .Values
                    .Cast<string>()
                    .ShouldHaveBeenSplitOnNewLinesAndEachLineInserted(combinedOutputXml);

                sourceFields
                    .Keys
                    .ShouldAll(k => combinedDocumentText.ShouldNotContain("«" + k + "»"));

            }
            finally{ output?.Dispose(); }
        }

        [SetUp]
        public void Setup() => sut = new MailMerger(Startup.Configure().CreateLogger(GetType()), Startup.Settings);

        [OneTimeSetUp]
        public void EnsureTestDependencies()
        {
            foreach(var testDoc in new[]{DocWithSimpleMergeFieldDocx})
            {
                var expectedTestDoc = Path.Combine(TestDocDir, testDoc);
                File.Exists(expectedTestDoc)
                    .ShouldBeTrue(
                        $"Expected to find TestDependency \n\n\"{expectedTestDoc}\"\n\n at "
                        + new FileInfo(expectedTestDoc).FullName + " but didn't. \n"
                        + "Include it in the test project and mark it as as BuildAction=Content, CopyToOutputDirectory=Copy if Newer."
                        );
            }
        }
    }
}
