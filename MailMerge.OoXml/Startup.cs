using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using MailMerge.OoXml.Properties;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace MailMerge.OoXml
{
    public static class Program
    {
        public static void Main(string[] args)
        {
            HelpAndExitIfNot( args.Length>0 );
            
            Startup.Configure(nameof(MailMerge));

            var component = new MailMerge(
                                           Startup.LoggerFactory.CreateLogger<MailMerge>(),
                                           Startup.Settings
                                          );

            var(files, mergefields) = ParseArgs.FromStringArray(args);

            foreach (var (filein,fileout) in files)
            {
                component.Merge(filein.Open(FileMode.Open,FileAccess.Read,FileShare.Read), mergefields, fileout.FullName );
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
            @"Help string for command line invocation here.

    Usage: MailMerge [fileName[...]] [ key=value[...] ]

    Settings can re read from the app-settings.json section {nameof(MailMerge)}

    Output is to StdOut.";

    }

    static class Startup
    {
        public static IConfiguration Configuration;
        public static ILoggerFactory LoggerFactory;
        public static Settings Settings;


        public static void Configure(string settingsName)
        {
            Configuration = new ConfigurationBuilder()
                           .SetBasePath(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location))
                           .AddJsonFile("appsettings.json", false)
                           .Build();
            Configuration.GetSection(settingsName).Bind(Settings = new Settings());

            LoggerFactory =  new LoggerFactory().FromConfiguration(Configuration);
        }
        
    }
}
