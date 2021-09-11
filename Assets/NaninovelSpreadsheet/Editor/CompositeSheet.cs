using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Naninovel.Parsing;
using UnityEngine;

namespace Naninovel.Spreadsheet
{
    public class CompositeSheet
    {
        private const string templateHeader = "Template";
        private const string argumentHeader = "Arguments";

        private readonly Dictionary<string, List<string>> columns = new Dictionary<string, List<string>>();
        private readonly Dictionary<int, string> localeTagsCache = new Dictionary<int, string>();
        private readonly List<Token> tokens = new List<Token>();
        private readonly Lexer lexer = new Lexer();

        public CompositeSheet (ScriptText script, IReadOnlyCollection<ScriptText> localizations)
        {
            FillColumnsFromScript(script, localizations);
        }

        public CompositeSheet (string managedText, IReadOnlyCollection<string> localizations)
        {
            FillColumnsFromManagedText(managedText, localizations);
        }

        public CompositeSheet (SpreadsheetDocument document, Worksheet sheet)
        {
            FillColumnsFromSpreadsheet(document, sheet);
        }

        public void WriteToSpreadsheet (SpreadsheetDocument document, Worksheet sheet)
        {
            for (int i = 0; i < columns.Count; i++)
            {
                var kv = columns.ElementAt(i);
                var header = kv.Key;
                var values = kv.Value;
                var rowNumber = (uint)1;
                var columnName = OpenXML.GetColumnNameFromNumber(i + 1);
                sheet.ClearAllCellsInColumn(columnName);
                document.SetCellValue(sheet, columnName, rowNumber, header);
                foreach (var value in values)
                {
                    rowNumber++;
                    document.SetCellValue(sheet, columnName, rowNumber, value);
                }
            }
            sheet.Save();
        }

        public void WriteToProject (string path, IReadOnlyCollection<string> localizations, bool managedText)
        {
            var builders = new Dictionary<string, StringBuilder>();
            var lastTemplateIndex = -1;
            var maxLength = columns.Values.Max(v => v.Count);
            for (int i = 0; i < maxLength; i++)
            {
                var template = GetColumnValues(templateHeader).ElementAtOrDefault(i);
                if (string.IsNullOrWhiteSpace(template)) continue;
                if (lastTemplateIndex > -1)
                    WriteLine(i - 1);
                lastTemplateIndex = i;
            }
            WriteLine(maxLength - 1, true);

            foreach (var kv in builders)
            {
                var header = kv.Key;
                var builder = kv.Value;
                if (header == argumentHeader)
                {
                    File.WriteAllText(path, GetBuilder(argumentHeader).ToString());
                    continue;
                }

                var localizationPath = localizations.FirstOrDefault(p => p.Contains($"/{header}/"));
                if (localizationPath is null || !File.Exists(localizationPath))
                    throw new Exception($"Localization document for `{header}` not found. Try re-generating the localization documents.");
                var localeHeader = File.ReadAllText(localizationPath).SplitByNewLine()[0];
                if (managedText) localeHeader += Environment.NewLine;
                File.WriteAllText(localizationPath, localeHeader + Environment.NewLine + builder);
            }

            void WriteLine (int lastArgIndex, bool lastLine = false)
            {
                var template = GetColumnValues(templateHeader)[lastTemplateIndex];
                var sourceArgs = GetColumnValuesAt(argumentHeader, lastTemplateIndex, lastArgIndex);
                foreach (var header in columns.Keys)
                {
                    if (header == templateHeader) continue;
                    var builder = GetBuilder(header);
                    if (header == argumentHeader)
                    {
                        builder.Append(new Composite(template, sourceArgs).Value);
                        if (lastLine || managedText) builder.AppendLine();
                        continue;
                    }

                    var localizedArgs = GetColumnValuesAt(header, lastTemplateIndex, lastArgIndex);
                    if (localizedArgs.Count > 0)
                        AppendLocalizationLine(builder, template, localizedArgs, sourceArgs);
                }
            }

            StringBuilder GetBuilder (string header)
            {
                if (!builders.TryGetValue(header, out var builder))
                {
                    builder = new StringBuilder();
                    builders[header] = builder;
                }
                return builder;
            }

            void AppendLocalizationLine (StringBuilder builder, string template, IReadOnlyList<string> localizedArgs, IReadOnlyList<string> sourceArgs)
            {
                var localizableTemplate = template.TrimEnd(StringUtils.NewLineChars).SplitByNewLine().Last();
                var localizedLine = new Composite(localizableTemplate, localizedArgs).Value;
                if (managedText)
                {
                    builder.AppendLine(localizedLine);
                    return;
                }
                var sourceLine = new Composite(localizableTemplate, sourceArgs).Value;
                var lineHash = CryptoUtils.PersistentHexCode(sourceLine.TrimFull());
                builder.AppendLine()
                    .AppendLine($"{Identifiers.LabelLine} {lineHash}")
                    .AppendLine($"{Identifiers.CommentLine} {sourceLine}");
                if (!localizedArgs.All(string.IsNullOrWhiteSpace))
                    builder.AppendLine(localizedLine);
            }
        }

        private void FillColumnsFromScript (ScriptText script, IReadOnlyCollection<ScriptText> localizations)
        {
            var templateBuilder = new StringBuilder();

            for (int i = 0; i < script.TextLines.Count; i++)
            {
                var lineText = script.TextLines[i];
                var line = script.Script.Lines[i];
                tokens.Clear();
                var lineType = lexer.TokenizeLine(lineText, tokens);
                var composite = new Composite(lineText, lineType, tokens);
                FillColumnsFromComposite(composite, templateBuilder);

                if (composite.Arguments.Count == 0) continue;
                foreach (var localizationScript in localizations)
                {
                    var locale = ExtractScriptLocaleTag(localizationScript);
                    var localizedValues = GetLocalizedValues(line, locale, localizationScript, composite.Arguments.Count);
                    GetColumnValues(locale).AddRange(localizedValues);
                }
            }

            if (templateBuilder.Length > 0)
            {
                var lastTemplateValue = templateBuilder.ToString().TrimEnd(StringUtils.NewLineChars);
                GetColumnValues(templateHeader).Add(lastTemplateValue);
            }

            string ExtractScriptLocaleTag (ScriptText localizationScript)
            {
                var firstCommentText = localizationScript.Script.Lines.OfType<CommentScriptLine>().FirstOrDefault()?.CommentText;
                return ExtractLocaleTag(localizationScript.GetHashCode(), firstCommentText);
            }

            IReadOnlyList<string> GetLocalizedValues (ScriptLine line, string locale, ScriptText localizationScript, int argsCount)
            {
                var startIndex = localizationScript.Script.GetLineIndexForLabel(line.LineHash);
                if (startIndex == -1)
                    throw new Exception($"Failed to find `{locale}` localization for `{script.Name}` script at line #{line.LineNumber}. Try re-generating localization documents.");
                var endIndex = localizationScript.Script.FindLine<LabelScriptLine>(l => l.LineIndex > startIndex)?.LineIndex ?? localizationScript.Script.Lines.Count;
                var localizationLines = localizationScript.Script.Lines
                    .Where(l => (l is CommandScriptLine || l is GenericTextScriptLine gl && gl.InlinedCommands.Count > 0) && l.LineIndex > startIndex && l.LineIndex < endIndex).ToArray();
                if (localizationLines.Length > 1)
                    Debug.LogWarning($"Multiple `{locale}` localization lines found for `{script.Name}` script at line #{line.LineNumber}. Only the first one will be exported to the spreadsheet.");
                if (localizationLines.Length == 0)
                    return Enumerable.Repeat(string.Empty, argsCount).ToArray();

                var localizationLineText = localizationScript.TextLines[localizationLines.First().LineIndex];
                tokens.Clear();
                var lineType = lexer.TokenizeLine(localizationLineText, tokens);
                var localizedComposite = new Composite(localizationLineText, lineType, tokens);
                if (localizedComposite.Arguments.Count != argsCount)
                    throw new Exception($"`{locale}` localization for `{script.Name}` script at line #{line.LineNumber} is invalid. Make sure it preserves original commands.");
                return localizedComposite.Arguments;
            }
        }

        private void FillColumnsFromManagedText (string managedText, IReadOnlyCollection<string> localizations)
        {
            var templateBuilder = new StringBuilder();
            var localizationLines = localizations.Select(l => l.SplitByNewLine()).ToArray();

            foreach (var line in managedText.SplitByNewLine())
            {
                if (!line.Contains(ManagedTextUtils.RecordIdLiteral) ||
                    line.StartsWithFast(ManagedTextUtils.RecordCommentLiteral)) continue;

                var composite = new Composite(line);
                FillColumnsFromComposite(composite, templateBuilder, false);

                if (composite.Arguments.Count == 0) continue;
                foreach (var localization in localizationLines)
                {
                    var locale = ExtractManagedTextLocaleTag(localization);
                    var localizedValue = GetLocalizedValue(line, locale, localization);
                    GetColumnValues(locale).Add(localizedValue);
                }
            }

            if (templateBuilder.Length > 0)
            {
                var lastTemplateValue = templateBuilder.ToString().TrimEnd(StringUtils.NewLineChars);
                GetColumnValues(templateHeader).Add(lastTemplateValue);
            }

            string ExtractManagedTextLocaleTag (string[] localization)
            {
                var firstCommentLine = localization.FirstOrDefault(l => l.StartsWithFast(ManagedTextUtils.RecordCommentLiteral));
                return ExtractLocaleTag(localization.GetHashCode(), firstCommentLine);
            }

            string GetLocalizedValue (string line, string locale, string[] localization)
            {
                var id = line.GetBefore(ManagedTextUtils.RecordIdLiteral);
                var localizedLine = localization.FirstOrDefault(l => l.StartsWithFast(id));
                if (localizedLine is null)
                {
                    Debug.LogWarning($"`{locale}` localization for `{id}` managed text is not found. Try re-generating localization documents.");
                    return string.Empty;
                }
                return localizedLine.Substring(id.Length + ManagedTextUtils.RecordIdLiteral.Length);
            }
        }

        private void FillColumnsFromSpreadsheet (SpreadsheetDocument document, Worksheet sheet)
        {
            for (int columnNumber = 1;; columnNumber++)
            {
                var columnName = OpenXML.GetColumnNameFromNumber(columnNumber);
                var cells = sheet.GetAllCellsInColumn(columnName)
                    .OrderBy(c => c.Ancestors<Row>().FirstOrDefault()?.RowIndex ?? uint.MaxValue).ToArray();
                if (cells.Length == 0) break;

                var header = cells[0].GetValue(document);
                if (columnNumber > 2 && !LanguageTags.ContainsTag(header)) break;

                for (int rowIndex = 1; rowIndex < cells.Length; rowIndex++)
                {
                    var cell = cells[rowIndex];
                    var cellValue = cell.GetValue(document);
                    GetColumnValues(header).Add(cellValue);
                }
            }
        }

        private void FillColumnsFromComposite (Composite composite, StringBuilder templateBuilder, bool appendLine = true)
        {
            templateBuilder.Append(composite.Template);
            if (appendLine) templateBuilder.AppendLine();
            if (composite.Arguments.Count == 0) return;

            foreach (var arg in composite.Arguments)
            {
                if (templateBuilder.Length > 0)
                {
                    GetColumnValues(templateHeader).Add(templateBuilder.ToString());
                    templateBuilder.Clear();
                }
                else GetColumnValues(templateHeader).Add(string.Empty);
                GetColumnValues(argumentHeader).Add(arg);
            }
        }

        private string ExtractLocaleTag (int cacheKey, string content)
        {
            if (localeTagsCache.TryGetValue(cacheKey, out var result))
                return result;

            var tag = content?.GetAfter("<")?.GetBefore(">");
            if (string.IsNullOrWhiteSpace(tag))
                throw new Exception($"Failed to extract localization tag from `{content}`. Try re-generating localization documents.");

            localeTagsCache[cacheKey] = tag;
            return tag;
        }

        private List<string> GetColumnValues (string header)
        {
            if (!columns.TryGetValue(header, out var values))
            {
                values = new List<string>();
                columns[header] = values;
            }
            return values;
        }

        private List<string> GetColumnValuesAt (string header, int startIndex, int endIndex)
        {
            var values = GetColumnValues(header);
            var length = Mathf.Min(values.Count - 1, endIndex) - startIndex + 1;
            return values.GetRange(startIndex, length);
        }
    }
}
