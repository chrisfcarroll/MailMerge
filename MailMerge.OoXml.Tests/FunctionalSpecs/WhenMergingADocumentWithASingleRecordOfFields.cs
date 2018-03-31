using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using MailMerge.Helpers;
using NUnit.Framework;
using TestBase;

namespace MailMerge.OoXml.Tests.FunctionalSpecs
{
    [TestFixture]
    public class WhenMergingADocumentWithASingleRecordOfFields
    {
        MailMerger sut;
        const string TemplateDocx = "TestDocuments\\ATemplate.docx";

        static Dictionary<string, string> MergeFieldsForTemplateDocx = new Dictionary<string, string>
        {
            {"FirstName","FakeFirst"},
            {"LastName","FakeLast"},
        };
        static Dictionary<string, string> MergeFieldsForClientCareLetter = new Dictionary<string, string>
        {
            {"Recipient.Salutation","Mr Test"},
            {"Matter.ClientsReference","TheClientsReference"},
            {"Matter.Reference","TestMatterRef-000-000-etc"},
            {"PropertyAddressOnOneLine","1, The Line, Property Address"},
            {"Sender","Test Letterwriter"},
        };

        [SetUp]
        public void Setup()
        {
            sut = new MailMerger(Startup.Configure().CreateLogger(GetType()), Startup.Settings);
        }

        [TestCase("TestDocuments\\ATemplate.docx", nameof(MergeFieldsForTemplateDocx))]
        [TestCase("TestDocuments\\Client Care Letter.docx", nameof(MergeFieldsForClientCareLetter))]
        public void Returns_TheDocumentWithMergeFieldsReplaced(string source, string sourceFieldsSource)
        {
            var sourceFields = GetType().GetField(sourceFieldsSource,BindingFlags.Static|BindingFlags.NonPublic).GetValue(this) as Dictionary<string,string>;

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


                Regex.IsMatch(outputText, @"\<w:instrText [^>]*>MERGEFIELD").ShouldBeFalse("Merge should have remove all complex field sequences (whether or not they were replaced with text).");

            }
            finally{ output?.Dispose(); }
        }

        [OneTimeSetUp]
        public void EnsureTestDependencies()
        {
            File.Exists(TemplateDocx).ShouldBeTrue($"TestDependency: \"{TemplateDocx}\" should be marked as CopyToOutputDirectory but didn't find it at {new FileInfo(TemplateDocx).FullName}");
        }

    }
}
