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
    public class GivenComplexMergeFields
    {
        const string PathToExampleDocs = "TestDocuments";
        const string ExampleDocx1 = "ATemplate.docx";
        const string ExampleDocx2 = "TestTemplate2.docx";
        MailMerger sut;

        static Dictionary<string, string> ExampleDocx1Fields = new Dictionary<string, string>
        {
            {"FirstName","FakeFirst"},
            {"LastName","FakeLast"},
        };
        static Dictionary<string, string> ExampleDocx2Fields = new Dictionary<string, string>
        {
            {"Recipient.Salutation","Mr Test"},
            {"Matter.ClientsReference","TheClientsReference"},
            {"Matter.Reference","TestMatterRef-000-000-etc"},
            {"PropertyAddressOnOneLine","1, The Line, Property Address"},
            {"Sender","Test Letterwriter"},
        };
        static Dictionary<string, string> ExampleDocx2FieldsWithMultiLines = new Dictionary<string, string>
        {
            {"Recipient.Salutation","Mr Test"},
            {"Matter.ClientsReference","TheClientsReference is surprisingly long \n" +
                                       "and has multiple lines \n" +
                                       "to it."},
            {"Matter.Reference","TestMatterRef-000-000-etc"},
            {"PropertyAddressOnOneLine","1, The First Line, \n" +
                                        "Is supposed to be one line,\n" +
                                        "But isn't\n" +
                                        "It's four lines."},
            {"Sender","Test Letterwriter"},
        };

        [SetUp]
        public void Setup()
        {
            sut = new MailMerger(Startup.Configure().CreateLogger(GetType()), Startup.Settings);
        }

        [TestCase(ExampleDocx1, nameof(ExampleDocx1Fields))]
        [TestCase(ExampleDocx2, nameof(ExampleDocx2Fields))]
        public void Returns_TheDocumentWithSingleLineMergeFieldsReplaced(string source, string sourceFieldsSource)
        {
            //A
            var sourceFields = GetType()
                .GetField(sourceFieldsSource, BindingFlags.Static | BindingFlags.NonPublic)
                .GetValue(this) as Dictionary<string, string>;
            //A
            (var outputText, var outputXml)=MergeDocToTextAndXml(source, sourceFields, " singleline Fields Output.");
            //A
            var singleLineSourceFields = sourceFields
                .Values
                .Where(v => v.DoesNotContain('\n'));
            singleLineSourceFields.ShouldAll(v => outputText.ShouldContain(v));
        }

        [TestCase(ExampleDocx1, nameof(ExampleDocx1Fields))]
        [TestCase(ExampleDocx2, nameof(ExampleDocx2Fields))]
        public void Treats_MergeFieldNamesAsCaseInsensitive(string source, string sourceFieldsSource)
        {
            //A
            var originalSourceFields = GetType()
                .GetField(sourceFieldsSource, BindingFlags.Static | BindingFlags.NonPublic)
                .GetValue(this) as Dictionary<string, string>;

            var sourceWithLowerCaseKeys = originalSourceFields
                .Select(kv => KeyValuePair.Create(kv.Key.ToLower(), kv.Value))
                .ToDictionary(kv=>kv.Key,kv=>kv.Value);
            var sourceWithUpperCaseKeys = originalSourceFields
                .Select(kv => KeyValuePair.Create(kv.Key.ToUpper(), kv.Value))
                .ToDictionary(kv=>kv.Key,kv=>kv.Value);
            
            foreach(var fieldDict in new []{sourceWithLowerCaseKeys,sourceWithUpperCaseKeys})
            {
                //A
                (var outputText, var outputXml)=MergeDocToTextAndXml(source, fieldDict, " singleline Fields Output.");
                //A
                var singleLineSourceFields = fieldDict
                    .Values
                    .Where(v => v.DoesNotContain('\n'));
                singleLineSourceFields.ShouldAll(v => outputText.ShouldContain(v));
            }
        }
        
        [TestCase(ExampleDocx2, nameof(ExampleDocx2FieldsWithMultiLines))]
        public void Returns_TheDocumentWithMultilineMergeFieldsReplaced(string source, string sourceFieldsSource)
        {
            //A
            var sourceFields = GetType()
                .GetField(sourceFieldsSource, BindingFlags.Static | BindingFlags.NonPublic)
                .GetValue(this) as Dictionary<string, string>;
            //A
            (var outputText, var outputXml)=MergeDocToTextAndXml(source, sourceFields, " Multiline Fields Output");
            //A
            var multiLineSourceFields = sourceFields
                .Values
                .Where(v => v.Contains('\n'));
            multiLineSourceFields.ShouldHaveBeenSplitOnNewLinesAndEachLineInserted(outputXml);
        }


        (string outputText, string outputXml) MergeDocToTextAndXml(string source, Dictionary<string, string> sourceFields, string outputFileSuffix)
        {
            source = Path.Combine(PathToExampleDocs, source);

            Stream output=null;
            AggregateException exceptions;

            using (var original = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read))
            try
            {
                (output, exceptions) = sut.Merge(original, sourceFields);

                using (var outFile = new FileInfo(source.Replace(".", outputFileSuffix+".")).OpenWrite())
                {
                    output.Position = 0;
                    output.CopyTo(outFile);
                }
                if (exceptions.InnerExceptions.Any()) { throw exceptions; }

                var (outputText, outputXml) = output.AsWordDocumentMainPartTextAndXml();
            
                sourceFields
                    .Keys
                    .ShouldAll(k => outputText.ShouldNotContain("«" + k + "»"));
                Regex
                    .IsMatch(outputText, @"\<w:instrText [^>]*>MERGEFIELD")
                    .ShouldBeFalse(
                        "Merge should have remove all complex field sequences " +
                        "(whether or not they were replaced with text).");
                return (outputText, outputXml);
            } 
            finally
            {
                output?.Dispose();                    
            }
        }


        [OneTimeSetUp]
        public void EnsureTestDependencies()
        {
            foreach(var testDoc in new[]{ExampleDocx1, ExampleDocx2})
            {
                var expectedTestDoc = Path.Combine(PathToExampleDocs, testDoc);
                File.Exists(expectedTestDoc)
                    .ShouldBeTrue(
                                  $"Expected to find TestDependency \n\n\"{expectedTestDoc}\"\n\n at "
                                + new FileInfo(expectedTestDoc).FullName + " but didn't. \n"
                                + "Include it in the test project and mark it as as BuildAction=Content, " 
                                + "CopyToOutputDirectory=Copy if Newer."
                                 );
            }
        }
    }
}
