using System;
using System.IO;
using MailMerge.Helpers;
using NUnit.Framework;
using TestBase;

namespace MailMerge.Tests.FunctionalSpecs
{
    [TestFixture]
    public class WhenUsingToolToExtractDocumentXml
    {
        const string PathToExampleDocs = "TestDocuments";
        const string ExampleDocx1 = "ATemplate.docx";
        const string ExampleDocx2 = "TestTemplate2.docx";
        const string KnownSnippetInDocx1 = @"This is ATemplate.docx";
        const string KnownSnippetInDocx2= @"Thank you for recently providing me with instructions to act for you in connection with the above transaction. I confirm that I shall be very happy to undertake this piece of work for you.";
        
        MailMerger sut;
        [SetUp]
        public void Setup() => 
            sut = new MailMerger(Startup.Configure().CreateLogger(GetType()), Startup.Settings);

        [TestCase(ExampleDocx1, KnownSnippetInDocx1)]
        [TestCase(ExampleDocx2, KnownSnippetInDocx2)]
        public void ItExtractsXmlDocument(string docxFileName, string knownSnippet)
        {
            var source = Path.Combine(PathToExampleDocs, docxFileName);
            var xml= new FileInfo(source).GetXmlDocumentOfWordprocessingMainDocument();
            Console.WriteLine(xml.OuterXml);
            xml.InnerText.ShouldContain(knownSnippet);
        }
        
        [TestCase(ExampleDocx1, KnownSnippetInDocx1)]
        [TestCase(ExampleDocx2, KnownSnippetInDocx2)]
        public void ItExtractsAnXElement(string docxFileName, string knownSnippet)
        {
            var source = Path.Combine(PathToExampleDocs, docxFileName);
            var xml= new FileInfo(source).GetXElementOfWordprocessingMainDocument();
            Console.WriteLine(xml.ToString());
            xml.ToString().ShouldContain(knownSnippet);
        }
    }
}