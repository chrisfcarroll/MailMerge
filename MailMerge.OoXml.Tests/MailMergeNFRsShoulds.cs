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
            sut.Merge(new MemoryStream(new byte[0]), new Dictionary<string,string>());
            Assert.That(logger.LoggedLines, Is.NotEmpty);
        }
        
        [Test]
        public void ReturnException__GivenNullInput()
        {
            var mergefields = new Dictionary<string, string>()
            {
                {"a", "aa"}
            };
            //
            var (errors,result)= new MailMerge(logger, new Settings()).Merge(null,mergefields);
            //
            errors.InnerExceptions.ShouldNotBeEmpty()[0].ShouldBeAssignableTo<ArgumentNullException>();
        }
    }
}
