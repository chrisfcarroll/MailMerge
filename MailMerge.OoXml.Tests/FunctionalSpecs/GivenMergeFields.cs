using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using System.Xml;
using System.Xml.XPath;
using MailMerge.OoXml.Properties;
using Microsoft.VisualStudio.TestPlatform.Utilities;
using NUnit.Framework;
using TestBase;
using Assert = TestBase.Assert;
using Is = TestBase.Is;

namespace MailMerge.OoXml.Tests
{
    [TestFixture]
    public class GivenMergeFields
    {
        MailMerge sut;
        StringListLogger logger;
        const string inputFile = "ATemplate.docx";

        Dictionary<string,string> MergeFieldsInTemplateDocx=new Dictionary<string, string>
        {
            {"FirstName","FakeFirst"},
            {"LastName","FakeLast"}
        };
        const int FileStreamInternalDefaultBufferSize=4096;

        [SetUp]public void Setup(){ sut = new MailMerge(logger = new StringListLogger(), new Settings()); }

        [Test]
        public void ShouldReturnOriginalDocumentWithMergeFieldsReplaced()
        {
            var (output, exceptions) = (Stream.Null, new AggregateException());
            using(var original = new FileStream(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            try
            {
                (output, exceptions) = sut.Merge(inputFile, new Dictionary<string, string>());
                //
                if (exceptions.InnerExceptions.Any()){ throw exceptions; }
                //
                foreach (var pair in MergeFieldsInTemplateDocx)
                {
                    output.Position = 0;
                    using (var reader = new StreamReader(output))
                    {
                        reader.ReadToEnd().ShouldContain(pair.Value).ShouldNotContain(pair.Key);
                    }
                }

                var ns= new XmlNamespaceManager(new NameTable());
                ns.AddNamespace("w","http://schemas.openxmlformats.org/wordprocessingml/2006/main");
                
                foreach(var pair in MergeFieldsInTemplateDocx)
                {
                    var nav = output.AsXPathDocOfWordprocessingMainDocument().CreateNavigator();

                    var foundKey=nav.Select($"@instr='{pair.Key}'", ns);
                    Console.WriteLine(foundKey);

                }
                var docAsXml = output.AsXPathDocOfWordprocessingMainDocument();

            }
            finally
            {
                output?.Dispose();
            }
        }

        [OneTimeSetUp]
        public void EnsureTestDependencies()
        {
            File.Exists(inputFile).ShouldBeTrue($"TestDependency: \"{inputFile}\" should be marked as CopyToOutputDirectory but didn't find it at {new FileInfo(inputFile).FullName}");
        }

    }
}
