using System;
using System.Collections.Generic;
using Naninovel.Parsing;

namespace Naninovel.Spreadsheet
{
    public class ScriptText
    {
        public Script Script { get; }
        public IReadOnlyList<string> TextLines { get; }
        public string Name => Script.Name;

        public ScriptText (Script script, string scriptText)
        {
            Script = script;
            TextLines = Helpers.SplitScriptText(scriptText);
            if (Script.Lines.Count != TextLines.Count)
                throw new Exception($"Failed to parse `{Script.name}` script: line count is not equal.");
        }
    }
}
