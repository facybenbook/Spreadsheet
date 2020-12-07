using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using UnityEngine;

namespace Naninovel.Spreadsheet
{
    public class CompositeSheet
    {
        private const string templateHeader = "Template";
        private const string argumentHeader = "Arguments";
        
        private readonly Dictionary<string, List<string>> columnValues = new Dictionary<string, List<string>>();
        private readonly Dictionary<int, string> localeTagsCache = new Dictionary<int, string>();
        private readonly StringBuilder templateBuilder = new StringBuilder();
        
        public CompositeSheet (Script script, IReadOnlyCollection<Script> localizations)
        {
            FillColumnsFromScript(script, localizations);
        }
        
        public CompositeSheet (string managedText, IReadOnlyCollection<string> localizations)
        {
            FillColumnsFromManagedText(managedText, localizations);
        }
        
        public void WriteToSpreadsheet (SpreadsheetDocument document, Worksheet sheet)
        {
            for (int i = 0; i < columnValues.Count; i++)
            {
                var kv = columnValues.ElementAt(i);
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

        private void FillColumnsFromScript (Script script, IReadOnlyCollection<Script> localizations)
        {
            foreach (var line in script.Lines)
            {
                var composite = new Composite(line);
                FillColumnsFromComposite(composite);
                
                if (composite.Arguments.Count == 0) continue;
                foreach (var localizationScript in localizations)
                {
                    var locale = ExtractScriptLocaleTag(localizationScript);
                    var localizedValues = GetLocalizedValues(line, locale, localizationScript, composite.Arguments.Count);
                    GetColumnValues(locale).AddRange(localizedValues);
                }
            }
            
            if (templateBuilder.Length > 0)
                GetColumnValues(templateHeader).Add(templateBuilder.ToString());
            
            string ExtractScriptLocaleTag (Script localizationScript)
            {
                var firstCommentText = localizationScript.Lines.OfType<CommentScriptLine>().FirstOrDefault()?.CommentText;
                return ExtractLocaleTag(localizationScript.GetHashCode(), firstCommentText);
            }
            
            IReadOnlyList<string> GetLocalizedValues (ScriptLine line, string locale, Script localizationScript, int argsCount)
            {
                var startIndex = localizationScript.GetLineIndexForLabel(line.LineHash);
                if (startIndex == -1)
                    throw new Exception($"Failed to find `{locale}` localization for `{line.ScriptName}` script at line #{line.LineNumber}. Try re-generating localization documents.");
                var endIndex = localizationScript.FindLine<LabelScriptLine>(l => l.LineIndex > startIndex)?.LineIndex ?? localizationScript.Lines.Count;
                var localizationLines = localizationScript.Lines
                    .Where(l => (l is CommandScriptLine || l is GenericTextScriptLine gl && gl.InlinedCommands.Count > 0) && l.LineIndex > startIndex && l.LineIndex < endIndex).ToArray();
                if (localizationLines.Length > 1)
                    Debug.LogWarning($"Multiple `{locale}` localization lines found for `{line.ScriptName}` script at line #{line.LineNumber}. Only the first one will be exported to the spreadsheet.");
                if (localizationLines.Length == 0)
                    return Enumerable.Repeat(string.Empty, argsCount).ToArray();
                
                var localizedComposite = new Composite(localizationLines.First());
                if (localizedComposite.Arguments.Count != argsCount)
                    throw new Exception($"`{locale}` localization for `{line.ScriptName}` script at line #{line.LineNumber} is invalid. Make sure it preserves original commands.");
                return localizedComposite.Arguments;
            }
        }

        private void FillColumnsFromManagedText (string managedText, IReadOnlyCollection<string> localizations)
        {
            var localizationLines = localizations.Select(l => l.SplitByNewLine()).ToArray();

            foreach (var line in managedText.SplitByNewLine())
            {
                if (!line.Contains(ManagedTextUtils.RecordIdLiteral) || 
                    line.StartsWithFast(ManagedTextUtils.RecordCommentLiteral)) continue;

                var composite = new Composite(line);
                FillColumnsFromComposite(composite);

                if (composite.Arguments.Count == 0) continue;
                foreach (var localization in localizationLines)
                {
                    var locale = ExtractManagedTextLocaleTag(localization);
                    var localizedValue = GetLocalizedValue(line, locale, localization);
                    GetColumnValues(locale).Add(localizedValue);
                }
            }

            if (templateBuilder.Length > 0)
                GetColumnValues(templateHeader).Add(templateBuilder.ToString());

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

        private void FillColumnsFromComposite (Composite composite)
        {
            templateBuilder.AppendLine(composite.Template);
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
            if (!columnValues.TryGetValue(header, out var values))
            {
                values = new List<string>();
                columnValues[header] = values;
            }
            return values;
        }
    }
}
