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
    public class GivenSplitWinstrFields
    {
        MailMerger sut;

        class TestCases
        {
            public const string TestDocDir = "TestDocuments";
            public const string DocWithSplitMergeFieldDocx = "DocWithSplitMergeField.docx";
            public const string DocProblem1Docx = "DocProblem1.docx";
            public const string DocWithWinstrTextDateRundocx = "DocWithWinstrTextDateRun.docx";
            public const string DocWithVerySplitMergeFielddocx = "DocWithVerySplitMergeField.docx";

            public static Dictionary<string, string> VerySplitMergeField = new()
            {
                { "SPLITMERGEFIELD", "Split Mergefield After Replacement" }
            };

            public static Dictionary<string, string> SplitField = new Dictionary<string, string>
            {
                {"CurrentUser:LastName","CurrentUserLastName"},
            };

            public static Dictionary<string, string> DocProblem1Fields = new Dictionary<string, string>
            {
                {"CurrentUser:FirstName","CurrentUserFirstName"}
            };

            public static Dictionary<string, string> DateOnly = new Dictionary<string, string>();
        }

        [SetUp]
        public void Setup()
        {
            sut = new MailMerger(Startup.Configure().CreateLogger(GetType()), Startup.Settings);
        }

        [TestCase(TestCases.DocWithWinstrTextDateRundocx, nameof(TestCases.DateOnly))]
        [TestCase(TestCases.DocWithSplitMergeFieldDocx, nameof(TestCases.SplitField))]
        [TestCase(TestCases.DocProblem1Docx, nameof(TestCases.DocProblem1Fields))]
        [TestCase(TestCases.DocWithVerySplitMergeFielddocx, nameof(TestCases.VerySplitMergeField) )]
        public void Returns_TheDocumentWithMergeFieldsReplaced(string source, string sourceFieldsSource)
        {
            source = Path.Combine(TestCases.TestDocDir, source);
            var sourceFields = typeof(TestCases)
                              .GetField(sourceFieldsSource,BindingFlags.Static|BindingFlags.Public)
                              .GetValue(this) as Dictionary<string,string>;

            Stream output = null; AggregateException exceptions;

            using (var original = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            try
            {
                (output, exceptions) = sut.Merge(original, sourceFields);

                using (var outFile = new FileInfo(source.Replace(".", " Output.")).OpenWrite())
                {
                    output.Position = 0;
                    output.CopyTo(outFile);
                }

                if (exceptions.InnerExceptions.Any()) { throw exceptions; }

                var outputText = output.AsWordprocessingDocument(false).MainDocumentPart.Document.InnerText;

                sourceFields
                    .Values
                    .ShouldAll(v => outputText.ShouldContain(v));

                sourceFields
                    .Keys
                    .ShouldAll(k => outputText.ShouldNotContain("«" + k + "»"));

                Regex
                    .IsMatch(outputText, @"\<w:instrText [^>]*>MERGEFIELD")
                    .ShouldBeFalse("Merge should have remove all complex field sequences (whether or not they were replaced with text).");

            }
            finally{ output?.Dispose(); }
        }

        [OneTimeSetUp]
        public void EnsureTestDependencies()
        {
            foreach(var testDoc in new[]{ TestCases.DocWithSplitMergeFieldDocx, TestCases.DocProblem1Docx})
            {
                var expectedTestDoc = Path.Combine(TestCases.TestDocDir, testDoc);
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
