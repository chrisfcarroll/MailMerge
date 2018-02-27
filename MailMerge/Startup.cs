using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using MailMerge.Properties;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Serilog;
using ILogger = Microsoft.Extensions.Logging.ILogger;

namespace MailMerge
{
    static class Startup
    {
        public static IConfiguration Configuration;
        public static ILoggerFactory LoggerFactory;
        public static Settings Settings;
        public static ILogger CreateLogger<T>() { return LoggerFactory.CreateLogger<T>(); }

        public static void Configure(string settingsName=null)
        {
            settingsName = settingsName ?? nameof(MailMerge);
            Configuration = new ConfigurationBuilder()
                           .SetBasePath(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location))
                           .AddJsonFile("appsettings.json", false)
                           .Build();
            Configuration.GetSection(settingsName).Bind(Settings = new Settings());

            LoggerFactory = new LoggerFactory().FromConfiguration(Configuration);
            LoggerFactory.CreateLogger("StartUp").LogDebug("Settings: {@Settings}",Settings);
        }
    }

    public static class Program
    {
        public static void Main(string[] args)
        {
            HelpAndExitIfNot( args.Length>0 );
            
            Startup.Configure();

            var component = new MailMerge(
                                           Startup.CreateLogger<MailMerge>(),
                                           Startup.Settings
                                          );

            var(files, mergefields) = ParseArgs.FromStringArray(args);

            foreach (var (filein,fileout) in files)
            {
                component.Merge(filein.FullName, mergefields, fileout.FullName );
            }
        }

        internal static class ParseArgs
        {
            enum OddEven {Even,Odd}
            public static ( (FileInfo,FileInfo)[], Dictionary<string, string>) FromStringArray(params string[] args)
            {
                var files = new List<(FileInfo, FileInfo)>(); 
                var mergefields=new Dictionary<string,string>();
                var oddeven = OddEven.Even;
                string lastin=null;
                foreach (var arg in args)
                {
                    if (arg.Contains("="))
                    {
                        var kv= arg.Split('=', 2);
                        mergefields.Add( kv[0], kv[1]);
                    }
                    else if(oddeven==OddEven.Even)
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

                return (files.ToArray(), mergefields);
            }
        }


        internal static void HelpAndExitIfNot(bool ok)
        {
            if (ok) return;
            Console.WriteLine(Help);
            Environment.Exit(0);
        }
        
        static readonly string Help = 
            @"
MailMerge inputFile1 outputFile1 [[inputFileN outputFileN]...] [ key=value[...] ]

    Settings can be read from the app-settings.json file.

    Example

    MailMerge input1.docx output1Bill.docx  FirstName=Bill  'LastName=O Reilly'

";

    }
}
