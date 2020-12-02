using System.Collections.Generic;
using System.Text;

namespace Naninovel.Spreadsheet
{
    public class CompositeSheet
    {
        public readonly IReadOnlyList<SheetColumn> Columns;

        public CompositeSheet (IEnumerable<Composite> composites)
        {
            var lineColumnValues = new List<string>();
            var textColumnValues = new List<string>();
            
            var lineBuilder = new StringBuilder();
            foreach (var composite in composites)
            {
                lineBuilder.AppendLine(composite.Template);
                if (composite.Args.Count == 0) continue;

                foreach (var arg in composite.Args)
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
            
            Columns = new []
            {
                new SheetColumn("Line", lineColumnValues),
                new SheetColumn("Text", textColumnValues)
            };
        }
    }
}
