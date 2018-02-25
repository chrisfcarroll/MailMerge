// Copied from :
// Decompiled with JetBrains decompiler
// Type: Microsoft.Extensions.Logging.StringList.StringListLoggerProvider
// Assembly: Microsoft.Extensions.Logging.StringList, Version=2.0.0.0, Culture=neutral, PublicKeyToken=adb9793829ddae60
// MVID: 2603C9C8-DD69-4A3E-883D-B34FC012F771
// Assembly location: /Users/chris/.nuget/packages/microsoft.extensions.logging.StringList/2.0.0/lib/netstandard2.0/Microsoft.Extensions.Logging.StringList.dll

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Console;

namespace MailMerge.OoXml.Tests
{
    public static class StringListLoggerFactoryExtension
    {
        public static ILoggerFactory AddStringListLogger(this ILoggerFactory factory, List<string> backingList = null,
                                                         string name = null)
        {
            if (factory == null) throw new ArgumentNullException(nameof(factory));
            factory.AddProvider(new StringListLoggerProvider(backingList, name));
            return factory;
        }
    }

    public class StringListLogger : ILogger
    {
        static readonly string LoglevelPadding = ": ";

        static readonly string MessagePadding =
            new string(' ', LogLevel.Information.ToString().Length + LoglevelPadding.Length);

        static readonly string NewLineWithMessagePadding = Environment.NewLine + MessagePadding;
        [ThreadStatic] static StringBuilder logBuilder;
        Func<string, LogLevel, bool> filter;

        public StringListLogger(List<String> backingList = null, string name = null,
                                Func<string, LogLevel, bool> filter = null, bool includeScopes = true)
        {
            Name = name     ?? String.Empty;
            Filter = filter ?? ((category, logLevel) => true);
            IncludeScopes = includeScopes;
            LoggedLines = backingList ?? new List<string>();
        }

        public List<string> LoggedLines { get; }

        public Func<string, LogLevel, bool> Filter
        {
            protected internal get => filter;
            set => filter = value ?? throw new ArgumentNullException(nameof(value));
        }

        public bool IncludeScopes { get; set; }

        public string Name { get; }

        public void Log<TState>(LogLevel logLevel, EventId eventId, TState state,
                                Exception exception,
                                Func<TState, Exception, string> formatter)
        {
            if (!IsEnabled(logLevel)) return;
            if (formatter == null) throw new ArgumentNullException(nameof(formatter));
            var message = formatter(state, exception);
            if (string.IsNullOrEmpty(message) && exception == null) return;
            WriteMessage(logLevel, Name, eventId.Id, message, exception);
        }

        public bool IsEnabled(LogLevel logLevel) { return Filter(Name, logLevel); }

        public IDisposable BeginScope<TState>(TState state)
        {
            if (state == null) throw new ArgumentNullException(nameof(state));
            return ConsoleLogScope.Push(Name, state);
        }

        public virtual void WriteMessage(LogLevel logLevel, string logName, int eventId, string message,
                                         Exception exception)
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
            if (!string.IsNullOrEmpty(message))
            {
                builder.Append(MessagePadding);
                var length = builder.Length;
                builder.AppendLine(message);
                builder.Replace(Environment.NewLine, NewLineWithMessagePadding, length, message.Length);
            }

            if (exception      != null) builder.AppendLine(exception.ToString());
            if (builder.Length > 0) LoggedLines.Add($"[{logLevel.ToString()}] {builder}");

            builder.Clear();
            if (builder.Capacity > 1024) builder.Capacity = 1024;
            logBuilder = builder;
        }

        void GetScopeInformation(StringBuilder builder)
        {
            var consoleLogScope = ConsoleLogScope.Current;
            var empty = string.Empty;
            var length = builder.Length;
            for (; consoleLogScope != null; consoleLogScope = consoleLogScope.Parent)
            {
                var str = length != builder.Length
                              ? string.Format("=> {0} ", consoleLogScope)
                              : string.Format("=> {0}", consoleLogScope);
                builder.Insert(length, str);
            }

            if (builder.Length <= length)
                return;
            builder.Insert(length, MessagePadding);
            builder.AppendLine();
        }
    }


    [ProviderAlias("StringList")]
    public class StringListLoggerProvider : ILoggerProvider, IDisposable
    {
        static readonly Func<string, LogLevel, bool> FalseFilter = (cat, level) => false;
        static readonly Func<string, LogLevel, bool> TrueFilter = (cat, level) => true;
        readonly Func<string, LogLevel, bool> filter;
        readonly bool includeScopes;

        public StringListLoggerProvider()
        {
            filter = TrueFilter;
            includeScopes = true;
        }

        public StringListLoggerProvider(List<String> backingList = null, string name = null,
                                        Func<string, LogLevel, bool> filter = null, bool includeScopes = true)
        {
            this.filter = filter ?? TrueFilter;
            this.includeScopes = includeScopes;
            if (name        != null && backingList == null) CreateLogger(name);
            if (backingList != null)
                Loggers.GetOrAdd(name ?? String.Empty,
                                 s => new StringListLogger(backingList, s, this.filter, this.includeScopes));
        }

        public StringListLoggerProvider(List<string> logger) { Loggers.TryAdd("", new StringListLogger(logger)); }

        public ConcurrentDictionary<string, StringListLogger> Loggers { get; } =
            new ConcurrentDictionary<string, StringListLogger>();

        public ILogger CreateLogger(string name)
        {
            return Loggers.GetOrAdd(name ?? String.Empty, CreateLoggerImplementation);
        }

        public void Dispose() { }

        StringListLogger CreateLoggerImplementation(string name)
        {
            return new StringListLogger(new List<string>(), name, GetFilter(name), includeScopes);
        }

        Func<string, LogLevel, bool> GetFilter(string name) { return filter ?? FalseFilter; }
    }
}
