using System;
using System.Collections.Generic;
using System.IO;
using System.IO.MemoryMappedFiles;
using System.Linq;
using MailMerge.OoXml.Properties;
using Microsoft.VisualStudio.TestPlatform.Utilities;
using NUnit.Framework;
using TestBase;
using Assert = TestBase.Assert;
using Is = TestBase.Is;

namespace MailMerge.OoXml.Tests
{
    [TestFixture]
    public class GivenNoMergeFields
    {
        MailMerge sut;
        StringListLogger logger;
        string inputFile = "ATemplate.docx";
        const int FileStreamInternalDefaultBufferSize=4096;

        [SetUp]public void Setup(){ sut = new MailMerge(logger = new StringListLogger(), new Settings()); }

        [Test]
        public void ShouldReturnOriginalDocument()
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
