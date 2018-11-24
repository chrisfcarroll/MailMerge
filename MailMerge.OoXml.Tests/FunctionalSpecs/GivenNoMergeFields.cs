using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MailMerge.Properties;
using NUnit.Framework;
using TestBase;

namespace MailMerge.OoXml.Tests.FunctionalSpecs
{
    [TestFixture]
    public class GivenNoMergeFields
    {
        MailMerger sut;
        string inputFile = Path.Combine("TestDocuments","ATemplate.docx");

        [SetUp]public void Setup(){ sut = new MailMerger(Startup.Configure().CreateLogger<GivenNoMergeFields>(), new Settings()); }

        [Test]
        public void Returns_OriginalDocument()
        {
            Stream output=null;
            AggregateException exceptions;
            using(var original = new FileStream(inputFile, FileMode.Open, FileAccess.Read, FileShare.Read))
            try
            {
                (output, exceptions) = sut.Merge(inputFile, new Dictionary<string, string>());
                //
                if (exceptions.InnerExceptions.Any()){ throw exceptions; }
                //
                output.ShouldEqualByStreamContent(original);
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
