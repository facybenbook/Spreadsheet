using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Naninovel.Spreadsheet
{
    public class CompositeSheet
    {
        public const string LinesHeader = "Script Lines";
        public const string TextHeader = "Localizable Text";
        
        public readonly IReadOnlyList<SheetColumn> Columns;

        public CompositeSheet (Script script)
        {
            var lineColumnValues = new List<string>();
            var textColumnValues = new List<string>();
            
            var lineBuilder = new StringBuilder();
            foreach (var line in script.Lines)
            {
                var composite = new Composite(line);
                lineBuilder.AppendLine(composite.Template);
                if (composite.Arguments.Count == 0) continue;

                foreach (var arg in composite.Arguments)
                {
                    if (lineBuilder.Length > 0)
                    {
                        lineColumnValues.Add(lineBuilder.ToString());
                        lineBuilder.Clear();
                    }
                    else lineColumnValues.Add(string.Empty);
                    textColumnValues.Add(arg);
                }
            }
            if (lineBuilder.Length > 0)
                lineColumnValues.Add(lineBuilder.ToString());
            
            Columns = new []
            {
                new SheetColumn(LinesHeader, lineColumnValues),
                new SheetColumn(TextHeader, textColumnValues)
            };
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
    }
}
