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
    public class MailMergeShouldReturnExceptions
    {
        MailMerge sut;
        StringListLogger logger;

        [SetUp]
        public void Setup()
        {
            sut = new MailMerge(logger = new StringListLogger(), new Settings());
        }        
    }
}
