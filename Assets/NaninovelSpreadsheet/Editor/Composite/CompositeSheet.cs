using System.Collections.Generic;
using System.Text;

namespace Naninovel.Spreadsheet
{
    public class CompositeSheet
    {
        public const string LinesHeader = "Script Lines";
        public const string TextHeader = "Localizable Text";
        
        public readonly IReadOnlyList<SheetColumn> Columns;

        public CompositeSheet (IEnumerable<Composite> composites)
        {
            var lineColumnValues = new List<string>();
            var textColumnValues = new List<string>();
            
            var lineBuilder = new StringBuilder();
            foreach (var composite in composites)
            {
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
    }
}
