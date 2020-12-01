using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using UnityEditor;
using UnityEngine;

namespace Naninovel.Spreadsheet
{
    public class SpreadsheetWindow : EditorWindow
    {
        protected string SpreadsheetPath { get => PlayerPrefs.GetString(GetPrefName()); set { PlayerPrefs.SetString(GetPrefName(), value); ValidatePaths(); } }
        protected string ScriptFolderPath { get => PlayerPrefs.GetString(GetPrefName()); set => PlayerPrefs.SetString(GetPrefName(), value); }
        protected string TextFolderPath { get => PlayerPrefs.GetString(GetPrefName()); set => PlayerPrefs.SetString(GetPrefName(), value); }
        protected string LocalizationFolderPath { get => PlayerPrefs.GetString(GetPrefName()); set => PlayerPrefs.SetString(GetPrefName(), value); }

        private static readonly GUIContent spreadsheetPathContent = new GUIContent("Spreadsheet", "The spreadsheet file (.xls or .xlsx).");
        private static readonly GUIContent scriptFolderPathContent = new GUIContent("Scripts", "Folder containing naninovel script files (optional).");
        private static readonly GUIContent textFolderPathContent = new GUIContent("Text", "Folder containing managed text files (optional).");
        private static readonly GUIContent localizationFolderPathContent = new GUIContent("Localization", "Folder containing localization resources (optional).");

        private bool pathsValid = false;
        
        [MenuItem("Naninovel/Tools/Spreadsheet")]
        private static void OpenWindow ()
        {
            var position = new Rect(100, 100, 500, 200);
            GetWindowWithRect<SpreadsheetWindow>(position, true, "Spreadsheet", true);
        }
        
        private static string GetPrefName ([CallerMemberName] string name = "") => $"Naninovel.{nameof(SpreadsheetWindow)}.{name}";

        private void ValidatePaths ()
        {
            pathsValid = File.Exists(SpreadsheetPath) && Path.GetExtension(SpreadsheetPath) == ".xls" || Path.GetExtension(SpreadsheetPath) == ".xlsx";
        }
        
        private void OnGUI ()
        {
            EditorGUILayout.LabelField("Naninovel Spreadsheet", EditorStyles.boldLabel);
            EditorGUILayout.LabelField("The tool to export/import scenario script, managed text and localization data to/from a spreadsheet.", EditorStyles.wordWrappedMiniLabel);

            EditorGUILayout.Space();

            using (new EditorGUILayout.HorizontalScope())
            {
                SpreadsheetPath = EditorGUILayout.TextField(spreadsheetPathContent, SpreadsheetPath);
                if (GUILayout.Button("Select", EditorStyles.miniButton, GUILayout.Width(65)))
                    SpreadsheetPath = EditorUtility.OpenFilePanel(spreadsheetPathContent.text, "", "xls,xlsx");
            }

            using (new EditorGUILayout.HorizontalScope())
            {
                ScriptFolderPath = EditorGUILayout.TextField(scriptFolderPathContent, ScriptFolderPath);
                if (GUILayout.Button("Select", EditorStyles.miniButton, GUILayout.Width(65)))
                    ScriptFolderPath = EditorUtility.OpenFolderPanel(scriptFolderPathContent.text, "", "");
            }

            using (new EditorGUILayout.HorizontalScope())
            {
                TextFolderPath = EditorGUILayout.TextField(textFolderPathContent, TextFolderPath);
                if (GUILayout.Button("Select", EditorStyles.miniButton, GUILayout.Width(65)))
                    TextFolderPath = EditorUtility.OpenFolderPanel(textFolderPathContent.text, "", "");
            }

            using (new EditorGUILayout.HorizontalScope())
            {
                LocalizationFolderPath = EditorGUILayout.TextField(localizationFolderPathContent, LocalizationFolderPath);
                if (GUILayout.Button("Select", EditorStyles.miniButton, GUILayout.Width(65)))
                    LocalizationFolderPath = EditorUtility.OpenFolderPanel(localizationFolderPathContent.text, "", "");
            }

            GUILayout.FlexibleSpace();

            if (pathsValid)
            {
                if (GUILayout.Button("Export", GUIStyles.NavigationButton))
                    try { Export(); } finally { EditorUtility.ClearProgressBar(); }
                
                if (GUILayout.Button("Import", GUIStyles.NavigationButton))
                    try { Import(); } finally { EditorUtility.ClearProgressBar(); }
            }
            else EditorGUILayout.HelpBox("Spreadsheet path is not valid; make sure it points to an existing .xls or .xlsx file.", MessageType.Error);

            EditorGUILayout.Space();
        }

        private void Export ()
        {
            if (!EditorUtility.DisplayDialog("Export data to the spreadsheet?",
                "Are you sure you want to export the scenario scripts, managed text and localization data to the spreadsheet?\n\nThe spreadsheet content will be overwritten, existing data could be lost. The effect of this action is permanent and can't be undone, so make sure to backup the spreadsheet file before confirming.\n\nIn case the spreadsheet is currently open in another program, close the program before proceeding.", "Export", "Cancel")) return;

            DisplayProgress(SpreadsheetPath, 0);
            
            var document = SpreadsheetDocument.Open(SpreadsheetPath, true);
            var scriptPaths = Directory.GetFiles(ScriptFolderPath, "*.nani", SearchOption.AllDirectories);
            for (int pathIdx = 0; pathIdx < scriptPaths.Length; pathIdx++)
            {
                var scriptPath = scriptPaths[pathIdx];
                var scriptName = Path.GetFileNameWithoutExtension(scriptPath);
                DisplayProgress(scriptName, pathIdx / (float)scriptPaths.Length);

                var assetPath = PathUtils.AbsoluteToAssetPath(scriptPath);
                var script = AssetDatabase.LoadAssetAtPath<Script>(assetPath);
                var scriptText = File.ReadAllText(scriptPath, Encoding.UTF8);
                var textLines = Script.SplitScriptText(scriptText);
                Debug.Assert(script.Lines.Count == textLines.Length);
                
                var sheetName = $"Scripts{scriptPath.Remove(ScriptFolderPath).Replace('\\', '>').Replace('/', '>').Remove(".nani")}";
                var sheet = document.GetOrAddSheet(sheetName);

                document.SetCellValue(sheet, "A", 1, "Line");
                document.SetCellValue(sheet, "B", 1, "Text");
                
                for (int lineIdx = 0; lineIdx < textLines.Length; lineIdx++)
                {
                    var textLine = textLines[lineIdx];
                    var line = script.Lines[lineIdx];
                    var rowIndex = (uint)lineIdx + 2;
                    document.SetCellValue(sheet, "A", rowIndex, textLine);
                    //Debug.Log(document.GetCellValue(sheet, "A", rowIndex));

                    var composite = new Composite(line, textLine);
                    Debug.Log($"{composite.Template} -----> {string.Join(",", composite.Args)}");
                }
                
                sheet.Save();
            }
            
            document.Dispose();

            void DisplayProgress (string file, float progress) => EditorUtility.DisplayProgressBar("Exporting To Spreadsheet", $"Processing `{file}`...", progress);
        }
        
        private void Import ()
        {
            if (!EditorUtility.DisplayDialog("Import data from the spreadsheet?",
                "Are you sure you want to import the spreadsheet data to this project?\n\nAffected scenario scripts, managed text and localization documents will be overwritten, existing data could be lost. The effect of this action is permanent and can't be undone, so make sure to backup the project before confirming.\n\nIn case the spreadsheet is currently open in another program, close the program before proceeding.", "Import", "Cancel")) return;
            
        }
    }
}
