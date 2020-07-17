using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using MailMerge.Properties;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using ILogger = Microsoft.Extensions.Logging.ILogger;

[assembly: InternalsVisibleTo("MailMerge.Tests")]
[assembly: Guid("d1c4ab83-c553-4e3b-8e75-c9e76498206b")]

namespace MailMerge
{
    class Startup
    {
        public static IConfiguration Configuration;
        public static ILoggerFactory LoggerFactory;
        public static Settings Settings;
        public static ILogger CreateLogger<T>() { return LoggerFactory.CreateLogger<T>(); }
        public static ILogger CreateLogger(Type type) { return LoggerFactory.CreateLogger(type); }
        public static ILogger CreateLogger(string name) { return LoggerFactory.CreateLogger(name); }

        public static Instance Configure(string settingsName=null)
        {
            //
            ILoggerProvider provider = null /*, add your preferred LoggerProvider() here */;
            ILoggerFactory factory = null /*Add your preferred ILoggerFactory here*/;
            //

            settingsName = settingsName ?? nameof(MailMerger);
            var startupLocation = Path.GetDirectoryName(typeof(Startup).Assembly.Location);
            if (File.Exists(Path.Combine(startupLocation, "appsettings.json")))
            {
                Configuration = new ConfigurationBuilder()
                    .SetBasePath(startupLocation)
                    .AddJsonFile("appsettings.json", false)
                    .Build();
                Configuration.GetSection(settingsName).Bind(Settings = new Settings());
            }
            else
            {
                Configuration = new ConfigurationBuilder().Build();
                Settings=new Settings();
            }
            LoggerFactory = factory.FromConfiguration(Configuration, provider);
            LoggerFactory.CreateLogger("StartUp").LogDebug("Settings: {@Settings}",Settings);
            return new Instance();
        }

        public class Instance : Startup
        {
            // ReSharper disable MemberHidesStaticFromOuterClass
            public new IConfiguration Configuration => Startup.Configuration;
            public new ILoggerFactory LoggerFactory => Startup.LoggerFactory;
            public new Settings Settings => Startup.Settings;
            public new ILogger CreateLogger<T>() => Startup.CreateLogger<T>();
            public new ILogger CreateLogger(string name)=>Startup.CreateLogger(name);
            public new ILogger CreateLogger(Type type) => Startup.CreateLogger(type);
        }
    }

    public static class Program
    {
        public static void Main(params string[] args)
        {
            HelpAndExitIfNot( args.Length>0 );
            
            Startup.Configure();

            var component = new MailMerger(
                                           Startup.CreateLogger<MailMerger>(),
                                           Startup.Settings
                                          );

            var(files, mergefields) = ParseArgs.FromStringArray(args);

            HelpAndExitIfNot(files.Length>0);

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
