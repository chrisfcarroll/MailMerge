using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace MailMerge.Properties
{
    class FallbackLoggerFactory: ILoggerFactory
    {
        static readonly FallbackLoggerProvider provider= new FallbackLoggerProvider();
        public ILogger CreateLogger(string categoryName){return provider.CreateLogger(categoryName);}
        public void AddProvider(ILoggerProvider provider){}
        public void Dispose(){provider?.Dispose();}
    }

    static class FallbackLoggerFactoryExtension
    {
        public static ILoggerFactory AddFallbackLogger(this ILoggerFactory factory, OutputTo consoleorlistofstring, List<string> backingList = null, string name = "ListOfString")
        {
            if (factory == null) throw new ArgumentNullException(nameof(factory));
            if (consoleorlistofstring == OutputTo.Console)
                throw new ArgumentException($"Can't specify {consoleorlistofstring} with a BackingList",nameof(backingList));

            factory.AddProvider(new FallbackLoggerProvider(backingList??new List<string>(), name));
            return factory;
        }
        public static ILoggerFactory AddFallbackLogger(this ILoggerFactory factory, OutputTo consoleorlistofstring, string name = "Console")
        {
            if (factory == null) throw new ArgumentNullException(nameof(factory));
            if (consoleorlistofstring == OutputTo.ListOfString)
                throw new ArgumentException($"Can't specify {consoleorlistofstring} unless you also specify a BackingList",nameof(consoleorlistofstring));
            factory.AddProvider(new FallbackLoggerProvider(name));
            return factory;
        }
        public enum OutputTo {Console,ListOfString}
    }

    class FallbackLogger : ILogger
    {
        public static ConcurrentDictionary<string, FallbackLogger> Loggers { get; } = new ConcurrentDictionary<string, FallbackLogger>();

        Action<string> Output;
        
        static readonly string LoglevelPadding = ": ";

        static readonly string MessagePadding = new string(' ', LogLevel.Information.ToString().Length + LoglevelPadding.Length);

        static readonly string NewLineWithMessagePadding = Environment.NewLine + MessagePadding;
        [ThreadStatic] static StringBuilder logBuilder;
        Func<string, LogLevel, bool> filter;
        static readonly JsonSerializerSettings ErrorSwallowingJsonSerializerSettings = new JsonSerializerSettings{Error = (o, e) => { }, ReferenceLoopHandling = ReferenceLoopHandling.Ignore};

        public FallbackLogger(List<string> backingList, string name=null, Func<string, LogLevel, bool> filter = null, bool includeScopes = true)
        {
            Name = name ?? String.Empty;
            Filter = filter ?? ((category, logLevel) => true);
            IncludeScopes = includeScopes;
            LoggedLines = backingList ?? new List<string>();
            Output = LoggedLines.Add;
        }
        public FallbackLogger(string name, Func<string, LogLevel, bool> filter = null, bool includeScopes = true)
        {
            Name = name ?? String.Empty;
            Filter = filter ?? ((category, logLevel) => true);
            IncludeScopes = includeScopes;
            Output = Console.WriteLine;
        }

        public List<string> LoggedLines { get; }

        public Func<string, LogLevel, bool> Filter
        {
            protected internal get => filter;
            set => filter = value ?? throw new ArgumentNullException(nameof(value));
        }

        public bool IncludeScopes { get; set; }

        public string Name { get; }

        ScopeStack Scopes {get;}= new ScopeStack();

        public class ScopeStack : Stack<(string, object)>, IDisposable
        {
            public void Dispose(){Pop();}
            public new ScopeStack Push((string,object) item){base.Push(item);return this;}
        }

        void SwallowError(EventArgs e){}

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
                                Exception exception,
                                Func<TState, Exception, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;
            if (formatter == null) throw new ArgumentNullException(nameof(formatter));
            string message;
            try
            {
                try
                {
                    message = formatter(state, exception);
                }
                catch (FormatException)
                {
                    /*
                     * https://github.com/aspnet/Logging/issues/767
                     */
                    var stateFields = state.GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance);
                    var values = stateFields.Select(f => f.GetValue(state)).ToArray();
                    var asJson = JsonConvert.SerializeObject(values, ErrorSwallowingJsonSerializerSettings);
                    message = $"{asJson} {exception}";
                }
            }
            catch (Exception e)
            {
                message = String.Format("error trying to log {0} {1}\nException: {2}", state?.GetType(), exception, e);
            }

            if (String.IsNullOrEmpty(message) && exception == null) return;
            WriteMessage(logLevel, Name, eventId.Id, message, exception);
        }

        public bool IsEnabled(LogLevel logLevel) { return Filter(Name, logLevel); }

        public IDisposable BeginScope<TState>(TState state)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            Scopes.Push((Name, state));
            return Scopes;
        }

        public virtual void WriteMessage(LogLevel logLevel, string logName, int eventId, string message, Exception exception)
        {
            var builder = logBuilder;
            logBuilder = null;
            if (builder == null) builder = new StringBuilder();
            builder.Append(LoglevelPadding);
            builder.Append(logName);
            builder.Append("[");
            builder.Append(eventId);
            builder.AppendLine("]");
            if (IncludeScopes) GetScopeInformation(builder);
            if (!String.IsNullOrEmpty(message))
            {
                builder.Append(MessagePadding);
                var length = builder.Length;
                builder.AppendLine(message);
                builder.Replace(Environment.NewLine, NewLineWithMessagePadding, length, message.Length);
            }

            if (exception      != null) builder.AppendLine(exception.ToString());
            if (builder.Length > 0) Output($"[{logLevel.ToString()}] {builder}");

            builder.Clear();
            if (builder.Capacity > 1024) builder.Capacity = 1024;
            logBuilder = builder;
        }

        void GetScopeInformation(StringBuilder builder)
        {
            var length = builder.Length;
            foreach(var scope in Scopes)
            {
                var asString = scope.Item2 is Type t ? t.Name : scope.Item2;
                var str = length != builder.Length
                              ? String.Format("=> {0} ", asString)
                              : String.Format("=> {0}",  asString);
                builder.Insert(length, str);
            }

            if (builder.Length <= length)return;
            builder.Insert(length, MessagePadding);
            builder.AppendLine();
        }
    }


    class FallbackLoggerProvider : ILoggerProvider, IDisposable
    {
        static readonly Func<string, LogLevel, bool> FalseFilter = (cat, level) => false;
        static readonly Func<string, LogLevel, bool> TrueFilter = (cat, level) => true;
        readonly Func<string, LogLevel, bool> filter;
        readonly bool includeScopes;

        public FallbackLoggerProvider()
        {
            filter = TrueFilter;
            includeScopes = true;
        }

        public FallbackLoggerProvider(List<String> backingList, string name = null, Func<string, LogLevel, bool> filter = null, bool includeScopes = true)
        {
            this.filter = filter ?? TrueFilter;
            this.includeScopes = includeScopes;
            if (name!= null && backingList == null) CreateLogger(name);
            FallbackLogger.Loggers.GetOrAdd(name ?? String.Empty, n => new FallbackLogger(backingList, n, this.filter, this.includeScopes));
        }

        public FallbackLoggerProvider(string name ) { FallbackLogger.Loggers.TryAdd(name, new FallbackLogger("name")); }

        public ILogger CreateLogger(string name)
        {
            return FallbackLogger.Loggers.GetOrAdd(name ?? String.Empty, CreateLoggerImplementation);
        }

        public void Dispose()
        {
            FallbackLogger.Loggers.Clear();
        }

        FallbackLogger CreateLoggerImplementation(string name)
        {
            return new FallbackLogger(name, GetFilter(name), includeScopes);
        }

        Func<string, LogLevel, bool> GetFilter(string name) { return filter ?? FalseFilter; }
    }
}
