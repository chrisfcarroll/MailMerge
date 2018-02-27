using System;
using System.Collections.Generic;
using System.IO;
using MailMerge.OoXml.Properties;
using NUnit.Framework;
using TestBase;
using Assert = TestBase.Assert;
using Is = TestBase.Is;

namespace MailMerge.OoXml.Tests
{
    public class MailMergeNFRsShoulds
    {
        MailMerge sut;
        StringListLogger logger;

        [SetUp]
        public void Setup()
        {
            sut = new MailMerge(logger = new StringListLogger(), new Settings());
        }

        [Test]
        public void Log()
        {
            sut.Merge(new MemoryStream(new byte[0]), new Dictionary<string, string>());
            Assert.That(logger.LoggedLines, Is.NotEmpty);
        }

        [Test]
        public void ReturnException__GivenNullStreamInput()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerge(logger, new Settings()).Merge(null as FileStream, mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }

        [Test]
        public void ReturnException__GivenNullInputFile()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerge(logger, new Settings()).Merge(null as string, mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }

        [Test]
        public void ReturnException__GivenNullInput__GivenOutputFilepath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };

            var (result, errors) = new MailMerge(logger, new Settings()).Merge(null as Stream, mergefields, "");
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }

        [Test]
        public void ReturnException__GivenEmptyInputFilePath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerge(logger, new Settings()).Merge("", mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty();
        }

        [Test]
        public void ReturnException__GivenInvalidInputFilePath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerge(logger, new Settings()).Merge(" ", mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty();
        }

        [Test]
        public void ReturnTwoException__GivenInvalidInputFilePathAndInvalidOutputFilePath()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (result, errors) = new MailMerge(logger, new Settings()).Merge("", mergefields, "");
            //
            errors.InnerExceptions.ShouldBeOfLength(2);
        }
    }
}
