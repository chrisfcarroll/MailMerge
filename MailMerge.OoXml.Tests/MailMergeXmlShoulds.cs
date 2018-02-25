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
    public class MailMergeXmlShoulds
    {
        MailMerge sut;
        StringListLogger logger;

        [SetUp]
        public void Setup()
        {
            logger = new StringListLogger();
            sut = new MailMerge(logger, new Settings());
        }

        [Test]
        public void ParseInputWordDocAsXml()
        {

        }
        
    }
}
