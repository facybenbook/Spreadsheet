using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Naninovel.Commands;
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
        private static Regex argRegex = new Regex(@"(?<!\\)\{(\d+?)(?<!\\)\}");

        public Composite (string template, IEnumerable<string> args)
        {
            Template = template;
            Arguments = args?.ToArray() ?? emptyArgs;
            Value = BuildTemplate(Template, Arguments);
        }

        public Composite (ScriptLine scriptLine)
        {
            (Template, Arguments) = ParseScriptLine(scriptLine);
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

        private static (string template, IReadOnlyList<string> args) ParseScriptLine (ScriptLine line)
        {
            if (line is CommandScriptLine commandLine)
                return ParseCommandLine(commandLine);
            if (line is GenericTextScriptLine genericLine)
                return ParseGenericLine(genericLine);
            if (line is CommentScriptLine commentLine)
                return ($"{Constants.CommentLineId} {commentLine.CommentText}", emptyArgs);
            if (line is LabelScriptLine labelLine)
                return ($"{Constants.LabelLineId} {labelLine.LabelText}", emptyArgs);
            if (line is EmptyScriptLine emptyLine)
                return (string.Empty, emptyArgs);
            throw new Exception($"Unknown command line type: {line.GetType().Name}");
        }

        private static (string template, IReadOnlyList<string> args) ParseCommandLine (CommandScriptLine commandLine)
        {
            var (commandTemplate, args) = ParseCommand(commandLine.Command);
            return (Constants.CommandLineId + commandTemplate, args);
        }

        private static (string template, IReadOnlyList<string> args) ParseCommand (Command command, int argOffset = 0)
        {
            var args = new List<string>();
            var commandName = command.GetType().GetCustomAttribute<Command.CommandAliasAttribute>()?.Alias ?? command.GetType().Name.FirstToLower();
            var templateBuilder = new StringBuilder(commandName);
            var parameterFields = command.GetType().GetFields()
                .Where(f => typeof(ICommandParameter).IsAssignableFrom(f.FieldType));
            foreach (var field in parameterFields)
            {
                var parameter = field.GetValue(command) as ICommandParameter;
                if (parameter is null || !Command.Assigned(parameter)) continue;

                var value = Helpers.EncodeValue(parameter.ToString(), true, true);
                if (field.GetCustomAttribute<Command.ParameterDefaultValueAttribute>()?.Value == value) continue;

                templateBuilder.Append(" ");

                var name = field.GetCustomAttribute<Command.ParameterAliasAttribute>()?.Alias ?? field.Name.FirstToLower();
                if (name != Command.NamelessParameterAlias)
                    templateBuilder.Append(name).Append(Constants.ParamAssignId);

                if (Attribute.IsDefined(field, typeof(Command.LocalizableParameterAttribute)))
                {
                    args.Add(value);
                    templateBuilder.Append(BuildPlaceholder(args.Count - 1 + argOffset));
                }
                else templateBuilder.Append(value);
            }

            return (templateBuilder.ToString(), args);
        }

        private static (string template, IReadOnlyList<string> args) ParseGenericLine (GenericTextScriptLine line)
        {
            var args = new List<string>();
            var templateBuilder = new StringBuilder();

            var authorId = line.InlinedCommands.OfType<PrintText>()
                .FirstOrDefault(p => Command.Assigned(p.AuthorId) && !p.AuthorId.DynamicValue)?.AuthorId.Value;
            if (!string.IsNullOrEmpty(authorId))
            {
                templateBuilder.Append(authorId);
                var appearance = line.InlinedCommands.First() is ModifyCharacter mc
                                 && Command.Assigned(mc.IdAndAppearance) && !mc.IdAndAppearance.DynamicValue
                                 && mc.IdAndAppearance.NamedValue.HasValue
                    ? mc.IdAndAppearance.NamedValue
                    : null;
                if (!string.IsNullOrEmpty(appearance))
                    templateBuilder.Append(Constants.AuthorAppearanceId).Append(appearance);
                templateBuilder.Append(Constants.AuthorAssignId);
            }

            for (int i = 0; i < line.InlinedCommands.Count; i++)
            {
                var command = line.InlinedCommands[i];
                if (i == 0 && command is ModifyCharacter) continue;

                if (command is PrintText print)
                {
                    args.Add(print.Text.ToString());
                    templateBuilder.Append(BuildPlaceholder(args.Count - 1));
                    if (print.WaitForInput && i < line.InlinedCommands.Count - 1)
                        templateBuilder.Append("[i]");
                    continue;
                }

                var (commandTemplate, commandArgs) = ParseCommand(command, args.Count);
                args.AddRange(commandArgs);
                templateBuilder.Append($"[{commandTemplate}]");
            }

            return (templateBuilder.ToString(), args);
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
