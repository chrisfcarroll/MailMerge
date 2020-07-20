using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MailMerge.Helpers;
using Microsoft.Extensions.Logging;

namespace MailMerge.CommandLine
{
    public static class Program
    {
        public enum Command {Merge,ShowXml}

        public static void Main(params string[] args)
        {
            Help.PrintHelpAndExitIfNot( args.Length>0 );
            
            Startup.Configure();

            var(command, files, mergefields) = ParseArgs.FromStringArray(args);

            switch (command)
            {
                case Command.ShowXml:
                    Help.PrintHelpAndExitIfNot(files.Length>0);
                    ShowEachFileAsXml(files.Select(f => f.Item1).ToArray());
                    break;
                
                case Command.Merge:
                default:
                    Help.PrintHelpAndExitIfNot(files.Length>0);
                    Help.PrintHelpAndExitIfNot(mergefields.Count>0);
                    MergeEachInputToOutput(files, mergefields);
                    break;
            }
        }

        static void ShowEachFileAsXml(FileInfo[] files)
        {
            if (!files.First().Exists)
            {
                var message = $"Called with --showxml {files.First().FullName} but file not found";
                Startup.CreateLogger("Main").LogError(message);
                Console.WriteLine(message);
                Environment.Exit(1);
            }

            foreach (var fileInfo in files.Where(f => f.Exists))
            {
                Console.WriteLine(fileInfo.GetXmlDocumentOfWordprocessingMainDocument().OuterXml);
            }
        }

        static void MergeEachInputToOutput((FileInfo, FileInfo)[] files, Dictionary<string, string> mergefields)
        {
            var component = new MailMerger(
                Startup.CreateLogger<MailMerger>(),
                Startup.Settings
            );
            foreach (var (filein, fileout) in files)
            {
                component.Merge(filein.FullName, mergefields, fileout.FullName);
            }
        }
    }
}