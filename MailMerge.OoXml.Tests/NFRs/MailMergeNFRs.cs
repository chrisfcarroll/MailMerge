using System;
using System.Collections.Generic;
using System.IO;
using Extensions.Logging.ListOfString;
using MailMerge.Properties;
using NUnit.Framework;
using TestBase;
using Assert = TestBase.Assert;
using Is = TestBase.Is;

namespace MailMerge.OoXml.Tests.NFRs
{
    public class MailMergeNFRs
    {
        MailMerger sut;
        StringListLogger logger;

        [SetUp]
        public void Setup()
        {
            sut = new MailMerger(logger = new StringListLogger(), new Settings());
        }

        [Test]
        public void Logs_ToSpecifiedLogger()
        {
            sut.Merge(new MemoryStream(new byte[0]), new Dictionary<string, string>());
            Assert.That(logger.LoggedLines, Is.NotEmpty);
        }

        [Test]
        public void ReturnsException__GivenNullStreamInput()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerger(logger, new Settings()).Merge(null as FileStream, mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }

        [Test]
        public void ReturnsException__GivenNullInputFile()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerger(logger, new Settings()).Merge(null as string, mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }

        [Test]
        public void ReturnsException__GivenNullInput__GivenOutputFilepath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };

            var (result, errors) = new MailMerger(logger, new Settings()).Merge(null as Stream, mergefields, "");
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }

        [Test]
        public void ReturnsException__GivenEmptyInputFilePath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerger(logger, new Settings()).Merge("", mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty();
        }

        [Test]
        public void ReturnsException__GivenInvalidInputFilePath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerger(logger, new Settings()).Merge(" ", mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty();
        }

        [Test]
        public void ReturnsTwoExceptions__GivenInvalidInputFilePathAndInvalidOutputFilePath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerger(logger, new Settings()).Merge("", mergefields, "");
            //
            errors.InnerExceptions.ShouldBeOfLength(2);
        }
    }
}
