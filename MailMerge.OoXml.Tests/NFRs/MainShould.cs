using NUnit.Framework;
using TestBase;

namespace MailMerge.OoXml.Tests.NFRs
{
    public class MainShould
    {
        [TestCase("in.docx", "out.docx", "a=b", "aa=bb")]
        [TestCase("in.docx", "out.docx", "in2.docx", "out2.docx", "a=b", "aa=bb")]
        public static void ParseArgs(params string[] args)
        {
            var(files,mergefields) = Program.ParseArgs.FromStringArray(args);

            files[0].Item1.Name.ShouldBe(args[0]);
            files[0].Item2.Name.ShouldBe(args[1]);
            mergefields["a"].ShouldNotBeNull().ShouldBe("b");
            mergefields["aa"].ShouldNotBeNull().ShouldBe("bb");
            
        }
    }
}
