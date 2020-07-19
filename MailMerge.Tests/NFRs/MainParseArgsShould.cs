using System.Linq;
using NUnit.Framework;
using TestBase;

namespace MailMerge.Tests.NFRs
{
    public class MainParseArgsShould
    {
        [TestCase("in.docx", "out.docx", "a=b", "aa=bb")]
        [TestCase("in.docx", "out.docx", "in2.docx", "out2.docx", "a=b", "aa=bb")]
        [TestCase("in.docx", "out.docx", "a=b", "aa=bb")]
        [TestCase("--merge","in.docx", "out.docx", "in2.docx", "out2.docx", "a=b", "aa=bb")]
        [TestCase("-merge","in.docx", "out.docx", "in2.docx", "out2.docx", "a=b", "aa=bb")]
        public void ParseArgsForMerge(params string[] args)
        {
            var(command, files,mergefields) = Program.ParseArgs.FromStringArray(args);
            command.ShouldBe(Program.Command.Merge);
            files[0].Item1.Name.ShouldBe(args[0]);
            files[0].Item2.Name.ShouldBe(args[1]);
            mergefields["a"].ShouldNotBeNull().ShouldBe("b");
            mergefields["aa"].ShouldNotBeNull().ShouldBe("bb");
            
        }

        [TestCase("--showxml","in1.docx", "in2.docx", "in3.docx")]
        [TestCase("-ShowXml","in1.docx", "in2.docx", "in3.docx", "a=b", "aa=bb")]
        public void ParseArgsForShowXml(params string[] args)
        {
            var(command, files,mergefields) = Program.ParseArgs.FromStringArray(args);
            command.ShouldBe(Program.Command.ShowXml);
            files[0].Item1.Name.ShouldBe(args[1]);
            files[0].Item2.Name.ShouldBe(args[1]);
            files[1].Item1.Name.ShouldBe(args[2]);
            files[1].Item2.Name.ShouldBe(args[2]);
            files[2].Item1.Name.ShouldBe(args[3]);
            files[2].Item2.Name.ShouldBe(args[3]);
            files.Length.ShouldBe(args.Length-1);
        }
    }
}