using System.Collections.Generic;
using System.Linq;

namespace Naninovel.Spreadsheet
{
    public class SheetColumn
    {
        public readonly string Header;
        public readonly IReadOnlyList<string> Values;

        private static readonly string[] emptyValues = new string[0];

        public SheetColumn (string header, IEnumerable<string> values)
        {
            Header = header;
            Values = values?.ToArray() ?? emptyValues;
        }
    }
}
