using System.Collections.Generic;

namespace Naninovel.Spreadsheet
{
    public class SheetColumn
    {
        public readonly string Id;
        public readonly IReadOnlyCollection<string> Values;

        public SheetColumn (string id, IReadOnlyCollection<string> values)
        {
            Id = id;
            Values = values;
        }
    }
}
