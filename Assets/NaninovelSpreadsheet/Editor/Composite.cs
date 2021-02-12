using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Naninovel.Lexing;
using Naninovel.Parsing;

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
        private static Regex argRegex = new Regex(@"(?<!\\)\{(\d+?)(?<!\\)\}", RegexOptions.Compiled);
        private static readonly ProjectMetadata projectMeta = new ProjectMetadata();

        private readonly Parsing.CommandLineParser commandLineParser = new Parsing.CommandLineParser();
        private readonly Parsing.GenericTextLineParser genericLineParser = new Parsing.GenericTextLineParser();
        private readonly List<string> errors = new List<string>();

        static Composite ()
        {
            IDEMetadataWindow.GenerateCommandsMetadata(projectMeta, Command.CommandTypes.Values,
                _ => (null, null, null), _ => null);
        }

        public Composite (string template, IEnumerable<string> args)
        {
            Template = template;
            Arguments = args?.ToArray() ?? emptyArgs;
            Value = BuildTemplate(Template, Arguments);
        }

        public Composite (string lineText, LineType lineType, IReadOnlyList<Token> tokens)
        {
            (Template, Arguments) = ParseScriptLine(lineText, lineType, tokens);
            Value = BuildTemplate(Template, Arguments);
        }

        public Composite (string managedTextLine)
        {
            Value = managedTextLine;
            (Template, Arguments) = ParseManagedText(managedTextLine);
        }

        private static string BuildPlaceholder (int index) => $"{{{index}}}";

        private static string BuildTemplate (string template, IReadOnlyList<string> args)
        {
            if (args.Count == 0) return template;

            foreach (var match in argRegex.Matches(template).Cast<Match>())
            {
                var index = int.Parse(match.Value.GetBetween("{", "}"));
                var arg = args[index];
                template = template.Replace(match.Value, arg);
            }

            return template;
        }

        private (string template, IReadOnlyList<string> args) ParseScriptLine (string lineText, LineType lineType, IReadOnlyList<Token> tokens)
        {
            errors.Clear();
            switch (lineType)
            {
                case LineType.Label:
                case LineType.Comment:
                    return (lineText, emptyArgs);
                case LineType.Command:
                    var commandLine = commandLineParser.Parse(lineText, tokens, errors);
                    if (errors.Count > 0) throw new Exception($"Line `{lineText}` failed to parse: {errors[0]}");
                    var commandResult = ParseCommandLine(commandLine, lineText);
                    commandLineParser.ReturnLine(commandLine);
                    return commandResult;
                case LineType.GenericText:
                    var genericLine = genericLineParser.Parse(lineText, tokens, errors);
                    if (errors.Count > 0) throw new Exception($"Line `{lineText}` failed to parse: {errors[0]}");
                    var genericResult = ParseGenericLine(genericLine, lineText);
                    genericLineParser.ReturnLine(genericLine);
                    return genericResult;
                default: return (string.Empty, emptyArgs);
            }
        }

        private static (string template, IReadOnlyList<string> args) ParseCommandLine (CommandLine model, string lineText)
        {
            var (commandTemplate, args) = ParseCommand(model.Command, lineText);
            return (Constants.CommandLineId + commandTemplate, args);
        }

        private static (string template, IReadOnlyList<string> args) ParseCommand (Parsing.Command command, string lineText, int argOffset = 0)
        {
            var commandMeta = projectMeta.commands.FirstOrDefault(c => (c.id?.EqualsFastIgnoreCase(command.Identifier) ?? false) ||
                                                                       (c.alias?.EqualsFastIgnoreCase(command.Identifier) ?? false));
            if (commandMeta is null) throw new Exception($"Unknown command: `{command.Identifier}`");
            if (!commandMeta.localizable) return (lineText, emptyArgs);

            var args = new List<string>();
            foreach (var parameter in command.Parameters)
            {
                var meta = commandMeta.@params.FirstOrDefault(c => (c.id?.EqualsFastIgnoreCase(parameter.Identifier) ?? false) ||
                                                                   (c.alias?.EqualsFastIgnoreCase(parameter.Identifier) ?? false));
                if (meta is null) throw new Exception($"Unknown parameter in `{command.Identifier}` command: `{parameter.Identifier}`");
                if (!meta.localizable) continue;
                args.Add(parameter.Value);
                var placeholder = BuildPlaceholder(args.Count - 1 + argOffset);
                lineText = lineText.Remove(parameter.StartIndex, parameter.Length).Insert(parameter.StartIndex, placeholder);
            }

            return (lineText, args);
        }

        private static (string template, IReadOnlyList<string> args) ParseGenericLine (GenericTextLine model, string lineText)
        {
            var args = new List<string>();
            foreach (var content in model.Content)
                if (content is Parsing.Command command) ParseInlinedCommand(command);
                else ParseGenericText(content as GenericText);

            void ParseInlinedCommand (Parsing.Command command)
            {
                var (template, commandArgs) = ParseCommand(command, lineText, args.Count);
                lineText = template;
                args.AddRange(commandArgs);
            }

            void ParseGenericText (GenericText text)
            {
                var placeholder = BuildPlaceholder(args.Count - 1);
                lineText = lineText.Remove(text.StartIndex, text.Length).Insert(text.StartIndex, placeholder);
                args.Add(text.Text);
            }

            return (lineText, args);
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
