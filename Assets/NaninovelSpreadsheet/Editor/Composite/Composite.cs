using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Naninovel.Commands;

namespace Naninovel.Spreadsheet
{
    /// <summary>
    /// Represent a string template associated with arguments.
    /// </summary>
    public class Composite
    {
        public readonly string Value;
        public readonly string Template;
        public readonly IReadOnlyList<string> Arguments;

        private static readonly string[] emptyArgs = new string[0];

        public Composite (string template, IEnumerable<string> args)
        {
            Template = template;
            Arguments = args?.ToArray() ?? emptyArgs;
            Value = Arguments.Count > 0 ? string.Format(template, Arguments) : template;
        }
        
        public Composite (ScriptLine scriptLine, string scriptLineText)
        {
            Value = scriptLineText;
            (Template, Arguments) = ParseScriptLine(scriptLine, scriptLineText);
        }
        
        public Composite (string managedTextLine)
        {
            Value = managedTextLine;
            (Template, Arguments) = ParseManagedText(managedTextLine);
        }

        private static string BuildPlaceholder (int index) => $"{{{index}}}";

        private static (string template, IReadOnlyList<string> args) ParseScriptLine (ScriptLine line, string lineText)
        {
            if (line is CommandScriptLine commandLine && commandLine.Command is Command.ILocalizable) 
                return ParseCommand(commandLine.Command);
            if (line is GenericTextScriptLine) 
                return ParseGenericLine(lineText);
            return (lineText, emptyArgs);
        }
        
        private static (string template, IReadOnlyList<string> args) ParseCommand (Command command)
        {
            var args = new List<string>();
            var commandName = command.GetType().GetCustomAttribute<Command.CommandAliasAttribute>()?.Alias ?? command.GetType().Name.FirstToLower();
            var templateBuilder = new StringBuilder($"{CommandScriptLine.IdentifierLiteral}{commandName} ");
            var parameterFields = command.GetType().GetFields()
                .Where(f => typeof(ICommandParameter).IsAssignableFrom(f.FieldType));
            foreach (var field in parameterFields)
            {
                var parameter = field.GetValue(command) as ICommandParameter;
                if (parameter is null || !parameter.HasValue) continue;
                
                var name = field.GetCustomAttribute<Command.ParameterAliasAttribute>()?.Alias ?? field.Name.FirstToLower();
                if (name != Command.NamelessParameterAlias)
                    templateBuilder.Append(name).Append(Command.ParameterAssignLiteral);

                var value = Command.EscapeParameterValue(parameter.ToString());
                if (Attribute.IsDefined(field, typeof(Command.LocalizableParameterAttribute)))
                {
                    args.Add(value);
                    templateBuilder.Append(BuildPlaceholder(args.Count - 1)).Append(" ");
                }
                else templateBuilder.Append(value).Append(" ");
            }
            
            return (templateBuilder.ToString(), args);
        }
        
        private static (string template, IReadOnlyList<string> args) ParseGenericLine (string line)
        {
            var args = new List<string>();
            var templateBuilder = new StringBuilder();
            var lineBody = GenericTextScriptLine.ExtractAuthor(line, out _, out _);
            if (lineBody.Length < line.Length)
                templateBuilder.Append(line.Substring(0, line.Length - lineBody.Length));
            
            var inlinedMatches = GenericTextScriptLine.InlinedCommandRegex.Matches(lineBody).Cast<Match>();
            var prevMatchEndIndex = -1;
            foreach (var match in inlinedMatches)
            {
                var argIndex = prevMatchEndIndex + 1;
                var argLength = match.Index - argIndex;
                if (argLength > 0)
                    AppendBodyArg(argIndex, argLength);
                
                templateBuilder.Append(match.Value);
                prevMatchEndIndex = match.GetEndIndex();
            }

            var lastArgIndex = prevMatchEndIndex + 1;
            var lastArgLength = lineBody.Length - lastArgIndex;
            if (lastArgLength > 0)
                AppendBodyArg(lastArgIndex, lastArgLength);
            
            return (templateBuilder.ToString(), args);

            void AppendBodyArg (int bodyIndex, int argLength)
            {
                var arg = lineBody.Substring(bodyIndex, argLength);
                args.Add(arg);
                var placeholder = BuildPlaceholder(args.Count - 1);
                templateBuilder.Append(placeholder);
            }
        }
        
        private static (string template, IReadOnlyList<string> args) ParseManagedText (string line)
        {
            if (string.IsNullOrWhiteSpace(line) || !line.Contains(ManagedTextUtils.RecordIdLiteral))
                return (line, emptyArgs);
            
            var id = line.GetBefore(ManagedTextUtils.RecordIdLiteral);
            var lhsLength = id.Length + ManagedTextUtils.RecordIdLiteral.Length;
            var value = line.Substring(lhsLength);
            var template = line.Substring(0, lhsLength) + BuildPlaceholder(0);
            var args = new[] { value };
            return (template, args);
        }
    }
}
