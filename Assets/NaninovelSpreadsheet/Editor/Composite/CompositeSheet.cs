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
        public const string TemplateHeader = "Template";
        public const string ArgumentHeader = "Arguments";
        
        public readonly IReadOnlyList<SheetColumn> Columns;
        
        private static readonly Dictionary<Script, string> localesCache = new Dictionary<Script, string>();
        
        public CompositeSheet (Script script, IReadOnlyCollection<Script> localizationScripts)
        {
            Columns = ParseScript(script, localizationScripts);
        }

        public void WriteToSheet (SpreadsheetDocument document, Worksheet sheet)
        {
            for (int i = 0; i < Columns.Count; i++)
            {
                var column = Columns[i];
                var rowNumber = (uint)1;
                var columnName = OpenXML.GetColumnNameFromNumber(i + 1);
                sheet.ClearAllCellsInColumn(columnName);
                document.SetCellValue(sheet, columnName, rowNumber, column.Header);
                foreach (var value in column.Values)
                {
                    rowNumber++;
                    document.SetCellValue(sheet, columnName, rowNumber, value);
                }
            }
        }

        private static IReadOnlyList<SheetColumn> ParseScript (Script script, IReadOnlyCollection<Script> localizationScripts)
        {
            var templateValues = new List<string>();
            var argumentValues = new List<string>();
            var localizedValuesMap = new Dictionary<string, List<string>>();
            var templateBuilder = new StringBuilder();
            foreach (var line in script.Lines)
            {
                var composite = new Composite(line);
                ParseComposite(composite, templateBuilder, templateValues, argumentValues);
                
                foreach (var localizationScript in localizationScripts)
                {
                    var localizedValues = GetLocalizedValuesForLine(line, localizationScript, composite.Arguments.Count);
                    AddLocalizedValues(localizationScript, localizedValuesMap, localizedValues);
                }
            }
            if (templateBuilder.Length > 0)
                templateValues.Add(templateBuilder.ToString());
            
            var columns = new List<SheetColumn>
            {
                new SheetColumn(TemplateHeader, templateValues),
                new SheetColumn(ArgumentHeader, argumentValues)
            };
            
            foreach (var kv in localizedValuesMap)
            {
                var localizedColumn = new SheetColumn(kv.Key, kv.Value);
                columns.Add(localizedColumn);
            }

            return columns;
        }
        
        private static void AddLocalizedValues (Script localizationScript, IDictionary<string, List<string>> valuesMap, IEnumerable<string> values)
        {
            var locale = ExtractLocaleTag(localizationScript);
            if (!valuesMap.TryGetValue(locale, out var localizedValues))
            {
                localizedValues = new List<string>();
                valuesMap[locale] = localizedValues;
            }
            localizedValues.AddRange(values);
        }

        private static string ExtractLocaleTag (Script localizationScript)
        {
            if (localesCache.TryGetValue(localizationScript, out var result))
                return result;
            
            var commentLine = localizationScript.Lines.OfType<CommentScriptLine>().FirstOrDefault();
            var tag = commentLine?.CommentText?.GetAfter("<")?.GetBefore(">");
            if (string.IsNullOrWhiteSpace(tag))
                throw new Exception($"Failed to extract localization tag from `{localizationScript.Name}`. Try re-generating localization documents.");

            localesCache[localizationScript] = tag;
            return tag;
        }
        
        private static void ParseComposite (Composite composite, StringBuilder templateBuilder, IList<string> templateValues, IList<string> argumentValues)
        {
            templateBuilder.AppendLine(composite.Template);
            if (composite.Arguments.Count == 0) return;

            foreach (var arg in composite.Arguments)
            {
                if (templateBuilder.Length > 0)
                {
                    templateValues.Add(templateBuilder.ToString());
                    templateBuilder.Clear();
                }
                else templateValues.Add(string.Empty);
                argumentValues.Add(arg);
            }
        }

        private static IReadOnlyList<string> GetLocalizedValuesForLine (ScriptLine line, Script localizationScript, int argsCount)
        {
            var locale = ExtractLocaleTag(localizationScript);
            var startIndex = localizationScript.GetLineIndexForLabel(line.LineHash);
            if (startIndex == -1)
                throw new Exception($"Failed to find `{locale}` localization for `{line.ScriptName}` script at line #{line.LineNumber}. Try re-generating localization documents.");
            var endIndex = localizationScript.FindLine<LabelScriptLine>(l => l.LineIndex > startIndex)?.LineIndex ?? localizationScript.Lines.Count;
            var localizationLines = localizationScript.Lines
                .Where(l => (line is CommandScriptLine || line is GenericTextScriptLine genericLine && genericLine.InlinedCommands.Count > 0) && line.LineIndex > startIndex && line.LineIndex < endIndex).ToArray();
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
}
