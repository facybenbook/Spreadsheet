using System.Collections.Generic;

namespace Naninovel.Spreadsheet
{
    public class CompositeSheet
    {
        public readonly IReadOnlyCollection<SheetColumn> Columns;

        public CompositeSheet (IReadOnlyCollection<Composite> lines)
        {
            
        }
    }
}
