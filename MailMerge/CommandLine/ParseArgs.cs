using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailMerge.CommandLine
{
    static class ParseArgs
    {
        enum OddEven {Even,Odd}
        public static (Program.Command, (FileInfo,FileInfo)[], Dictionary<string, string>) FromStringArray(params string[] args)
        {
            var command = args.Length > 1 && args[0].StartsWith("-") 
                ? Enum.Parse<Program.Command>(args[0].TrimStart('-'),ignoreCase:true) 
                : Program.Command.Merge;
            var cargs= (command == Program.Command.Merge) ? args : args.Skip(1);

            switch (command)
            {
                case Program.Command.ShowXml :
                    return ParseArgsForShowXml(cargs, command);
                case Program.Command.Merge:
                default:
                    return ParseArgsForMerge(cargs, command);
            }
        }

        static (Program.Command, (FileInfo, FileInfo)[], Dictionary<string, string>) ParseArgsForShowXml(IEnumerable<string> cargs, Program.Command command)
        {
            return (command, cargs.Select(a => (new FileInfo(a), new FileInfo(a))).ToArray(), new Dictionary<string, string>());
        }

        static (Program.Command, (FileInfo, FileInfo)[], Dictionary<string, string>) ParseArgsForMerge(IEnumerable<string> cargs, Program.Command command)
        {
            var files = new List<(FileInfo, FileInfo)>();
            var mergefields = new Dictionary<string, string>(StringComparer.InvariantCultureIgnoreCase);
            var oddeven = OddEven.Even;
            string lastin = null;
            foreach (var arg in cargs)
            {
                if (arg.Contains("="))
                {
                    var kv = arg.Split('=', 2);
                    mergefields.Add(kv[0], kv[1]);
                }
                else if (oddeven == OddEven.Even)
                {
                    lastin = arg;
                    oddeven = OddEven.Odd;
                }
                else
                {
                    files.Add((new FileInfo(lastin), new FileInfo(arg)));
                    oddeven = OddEven.Even;
                }
            }

            return (command, files.ToArray(), mergefields);
        }
    }
}