// using System;
// using Naninovel.Spreadsheet;
// using UnityEngine;
//
// [SpreadsheetProcessor]
// public class TestCustomProcessor : SpreadsheetProcessor
// {
//     public TestCustomProcessor (Parameters parameters, Action<ProgressChangedArgs> onProgress)
//         : base(parameters, onProgress) { }
//
//     protected override ProjectWriter CreateProjectWriter (CompositeSheet composite)
//     {
//         Debug.Log("Override project writer (import from excel spreadsheet to script assets).");
//         return base.CreateProjectWriter(composite);
//     }
//
//     protected override SpreadsheetWriter CreateSpreadsheetWriter (CompositeSheet composite)
//     {
//         Debug.Log("Override spreadsheet writer (export from script assets to excel spreadsheet).");
//         return base.CreateSpreadsheetWriter(composite);
//     }
//
//     protected override ScriptReader CreateScriptReader (CompositeSheet composite)
//     {
//         Debug.Log("Override script reader (filling sheet columns from script assets).");
//         return base.CreateScriptReader(composite);
//     }
//
//     protected override ManagedTextReader CreateManagedTextReader (CompositeSheet composite)
//     {
//         Debug.Log("Override managed text reader (filling sheet columns from managed text docs).");
//         return base.CreateManagedTextReader(composite);
//     }
//
//     protected override SpreadsheetReader CreateSpreadsheetReader (CompositeSheet composite)
//     {
//         Debug.Log("Override spreadsheet reader (filling sheet columns from excel spreadsheet).");
//         return base.CreateSpreadsheetReader(composite);
//     }
// }
