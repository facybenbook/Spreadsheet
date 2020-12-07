using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using UnityEditor;
using UnityEngine;

namespace Naninovel.Spreadsheet
{
    public class SpreadsheetWindow : EditorWindow
    {
        protected string SpreadsheetPath { get => PlayerPrefs.GetString(GetPrefName()); set { PlayerPrefs.SetString(GetPrefName(), value); ValidatePaths(); } }
        protected string ScriptFolderPath { get => PlayerPrefs.GetString(GetPrefName()); set => PlayerPrefs.SetString(GetPrefName(), value); }
        protected string ManagedTextFolderPath { get => PlayerPrefs.GetString(GetPrefName()); set => PlayerPrefs.SetString(GetPrefName(), value); }
        protected string LocalizationFolderPath { get => PlayerPrefs.GetString(GetPrefName()); set => PlayerPrefs.SetString(GetPrefName(), value); }

        private static readonly GUIContent spreadsheetPathContent = new GUIContent("Spreadsheet", "The spreadsheet file (.xls or .xlsx).");
        private static readonly GUIContent scriptFolderPathContent = new GUIContent("Scripts", "Folder containing naninovel script files (optional).");
        private static readonly GUIContent textFolderPathContent = new GUIContent("Managed Text", "Folder containing managed text files (optional).");
        private static readonly GUIContent localizationFolderPathContent = new GUIContent("Localization", "Folder containing localization resources (optional).");

        private bool pathsValid = false;
        
        [MenuItem("Naninovel/Tools/Spreadsheet")]
        private static void OpenWindow ()
        {
            var position = new Rect(100, 100, 500, 200);
            GetWindowWithRect<SpreadsheetWindow>(position, true, "Spreadsheet", true);
        }
        
        private static string GetPrefName ([CallerMemberName] string name = "") => $"Naninovel.{nameof(SpreadsheetWindow)}.{name}";

        private void OnEnable ()
        {
            ValidatePaths();
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
                ManagedTextFolderPath = EditorGUILayout.TextField(textFolderPathContent, ManagedTextFolderPath);
                if (GUILayout.Button("Select", EditorStyles.miniButton, GUILayout.Width(65)))
                    ManagedTextFolderPath = EditorUtility.OpenFolderPanel(textFolderPathContent.text, "", "");
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
                    Export();
                
                if (GUILayout.Button("Import", GUIStyles.NavigationButton))
                    Import();
            }
            else EditorGUILayout.HelpBox("Spreadsheet path is not valid; make sure it points to an existing .xls or .xlsx file.", MessageType.Error);

            EditorGUILayout.Space();
        }
        
        private void ValidatePaths ()
        {
            pathsValid = File.Exists(SpreadsheetPath) && Path.GetExtension(SpreadsheetPath) == ".xls" || Path.GetExtension(SpreadsheetPath) == ".xlsx";
        }

        private void Export ()
        {
            if (!EditorUtility.DisplayDialog("Export data to the spreadsheet?",
                "Are you sure you want to export the scenario scripts, managed text and localization data to the spreadsheet?\n\nThe spreadsheet content will be overwritten, existing data could be lost. The effect of this action is permanent and can't be undone, so make sure to backup the spreadsheet file before confirming.\n\nIn case the spreadsheet is currently open in another program, close the program before proceeding.", "Export", "Cancel")) return;

            try
            {
                var document = SpreadsheetDocument.Open(SpreadsheetPath, true);
                ExportScriptsToSpreadsheet(document);
                ExportManagedTextToSpreadsheet(document);
                document.Dispose();
            }
            finally { EditorUtility.ClearProgressBar(); }
        }

        private void ExportScriptsToSpreadsheet (SpreadsheetDocument document)
        {
            if (!Directory.Exists(ScriptFolderPath)) return;
            
            var scriptPaths = Directory.GetFiles(ScriptFolderPath, "*.nani", SearchOption.AllDirectories);
            for (int pathIndex = 0; pathIndex < scriptPaths.Length; pathIndex++)
            {
                DisplayProgress("Exporting Naninovel Scripts", scriptPaths, pathIndex);
                var scriptPath = scriptPaths[pathIndex];
                var script = LoadScriptAtPath(scriptPath);
                var localPath = scriptPath.Remove(ScriptFolderPath).TrimStart('\\').TrimStart('/');
                var sheetName = $"Scripts>{localPath.Replace('\\', '>').Replace('/', '>').Remove(".nani")}";
                var sheet = document.GetSheet(sheetName) ?? document.AddSheet(sheetName);
                var localizations = LocateLocalizationsFor(localPath, ScriptsConfiguration.DefaultScriptsPathPrefix)
                    .Select(LoadScriptAtPath).ToArray();
                new CompositeSheet(script, localizations).WriteToSpreadsheet(document, sheet);
            }
        }
        
        private void ExportManagedTextToSpreadsheet (SpreadsheetDocument document)
        {
            if (!Directory.Exists(ManagedTextFolderPath)) return;
            
            var managedTextPaths = Directory.GetFiles(ManagedTextFolderPath, "*.txt", SearchOption.AllDirectories);
            for (int pathIndex = 0; pathIndex < managedTextPaths.Length; pathIndex++)
            {
                DisplayProgress("Exporting Managed Text", managedTextPaths, pathIndex);
                var docPath = managedTextPaths[pathIndex];
                var docText = File.ReadAllText(docPath, Encoding.UTF8);
                var localPath = docPath.Remove(ManagedTextFolderPath).TrimStart('\\').TrimStart('/');
                var sheetName = $"Text>{localPath.Replace('\\', '>').Replace('/', '>').Remove(".txt")}";
                var sheet = document.GetSheet(sheetName) ?? document.AddSheet(sheetName);
                var localizations = LocateLocalizationsFor(localPath, ManagedTextConfiguration.DefaultManagedTextPathPrefix)
                    .Select(p => File.ReadAllText(p, Encoding.UTF8)).ToArray();
                new CompositeSheet(docText, localizations).WriteToSpreadsheet(document, sheet);
            }
        }

        private void Import ()
        {
            if (!EditorUtility.DisplayDialog("Import data from the spreadsheet?",
                "Are you sure you want to import the spreadsheet data to this project?\n\nAffected scenario scripts, managed text and localization documents will be overwritten, existing data could be lost. The effect of this action is permanent and can't be undone, so make sure to backup the project before confirming.\n\nIn case the spreadsheet is currently open in another program, close the program before proceeding.", "Import", "Cancel")) return;
            
            try
            {

            }
            finally { EditorUtility.ClearProgressBar(); }
        }

        private static void DisplayProgress (string title, IReadOnlyList<string> paths, int index)
        {
            var path = paths[index];
            var name = Path.GetFileNameWithoutExtension(path);
            var progress = index / (float)paths.Count;
            EditorUtility.DisplayProgressBar(title, $"Processing `{name}`...", progress);
        }
        
        private static Script LoadScriptAtPath (string scriptPath)
        {
            var assetPath = PathUtils.AbsoluteToAssetPath(scriptPath);
            var script = AssetDatabase.LoadAssetAtPath<Script>(assetPath);
            if (script == null)
            {
                var scriptText = File.ReadAllText(scriptPath, Encoding.UTF8);
                script = Script.FromScriptText(scriptPath, scriptText);
            }
            if (script == null) throw new Exception($"Failed to load `{scriptPath}` script.");
            return script;
        }

        private IReadOnlyCollection<string> LocateLocalizationsFor (string localPath, string pathPrefix)
        {
            if (!Directory.Exists(LocalizationFolderPath)) return new string[0];
            
            var paths = new List<string>();
            foreach (var localeDir in Directory.EnumerateDirectories(LocalizationFolderPath))
            {
                var localizationPath = Path.Combine(localeDir, pathPrefix, localPath);
                if (!File.Exists(localizationPath))
                {
                    Debug.LogWarning($"Missing localization resource for `{localPath}` (expected in `{localizationPath}`)."); 
                    continue;
                }
                paths.Add(localizationPath);
            }
            return paths;
        }
    }
}
