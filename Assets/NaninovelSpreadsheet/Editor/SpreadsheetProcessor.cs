using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using UnityEditor;
using UnityEngine;

namespace Naninovel.Spreadsheet
{
    public class SpreadsheetProcessor
    {
        public class Parameters
        {
            public string SpreadsheetPath { get; set; }
            public string ScriptFolderPath { get; set; }
            public string ManagedTextFolderPath { get; set; }
            public string LocalizationFolderPath { get; set; }
        }

        private const string sheetPathSeparator = ">";
        private const string scriptSheetNamePrefix = "Scripts" + sheetPathSeparator;
        private const string textSheetNamePrefix = "Text" + sheetPathSeparator;
        private const string scriptFileExtension = ".nani";
        private const string textFileExtension = ".txt";
        private const string scriptFilePattern = "*" + scriptFileExtension;
        private const string textFilePattern = "*" + textFileExtension;

        private readonly string spreadsheetPath;
        private readonly string scriptFolderPath;
        private readonly string textFolderPath;
        private readonly string localeFolderPath;
        private readonly Action<ProgressChangedArgs> onProgress;
        
        public SpreadsheetProcessor (Parameters parameters, Action<ProgressChangedArgs> onProgress = default)
        {
            spreadsheetPath = parameters.SpreadsheetPath;
            scriptFolderPath = parameters.ScriptFolderPath;
            textFolderPath = parameters.ManagedTextFolderPath;
            localeFolderPath = parameters.LocalizationFolderPath;
            this.onProgress = onProgress;
        }

        public void Export ()
        {
            var document = SpreadsheetDocument.Open(spreadsheetPath, true);
            if (Directory.Exists(scriptFolderPath))
                ExportScriptsToSpreadsheet(document);
            if (Directory.Exists(textFolderPath))
                ExportManagedTextToSpreadsheet(document);
            document.Dispose();
        }

        public void Import ()
        {
            var document = SpreadsheetDocument.Open(spreadsheetPath, false);
            var sheetsNames = document.GetSheetNames().ToArray();
            for (int i = 0; i < sheetsNames.Length; i++)
            {
                NotifyProgressChanged(sheetsNames, i);
                var sheetName = sheetsNames[i];
                var sheet = document.GetSheet(sheetName);
                var localPath = SheetNameToLocalPath(sheetName);
                var fullPath = LocalToFullPath(localPath);
                var localizations = LocateLocalizationsFor(localPath, false);
                var compositeSheet = new CompositeSheet(document, sheet);
                var fromScript = fullPath.EndsWithFast(scriptFileExtension);
                if (fromScript) compositeSheet.WriteToScript(fullPath, localizations);
                else compositeSheet.WriteToManagedText(fullPath, localizations);
            }
            document.Dispose();
        }

        private void ExportScriptsToSpreadsheet (SpreadsheetDocument document)
        {
            var scriptPaths = Directory.GetFiles(scriptFolderPath, scriptFilePattern, SearchOption.AllDirectories);
            for (int pathIndex = 0; pathIndex < scriptPaths.Length; pathIndex++)
            {
                NotifyProgressChanged(scriptPaths, pathIndex);
                var scriptPath = scriptPaths[pathIndex];
                var script = LoadScriptAtPath(scriptPath);
                var localPath = FullToLocalPath(scriptPath);
                var sheetName = LocalPathToSheetName(localPath);
                var sheet = document.GetSheet(sheetName) ?? document.AddSheet(sheetName);
                var localizations = LocateLocalizationsFor(localPath)
                    .Select(LoadScriptAtPath).ToArray();
                new CompositeSheet(script, localizations).WriteToSpreadsheet(document, sheet);
            }
        }

        private void ExportManagedTextToSpreadsheet (SpreadsheetDocument document)
        {
            var managedTextPaths = Directory.GetFiles(textFolderPath, textFilePattern, SearchOption.AllDirectories);
            for (int pathIndex = 0; pathIndex < managedTextPaths.Length; pathIndex++)
            {
                NotifyProgressChanged(managedTextPaths, pathIndex);
                var docPath = managedTextPaths[pathIndex];
                var docText = File.ReadAllText(docPath, Encoding.UTF8);
                var localPath = FullToLocalPath(docPath);
                var sheetName = LocalPathToSheetName(localPath);
                var sheet = document.GetSheet(sheetName) ?? document.AddSheet(sheetName);
                var localizations = LocateLocalizationsFor(localPath)
                    .Select(p => File.ReadAllText(p, Encoding.UTF8)).ToArray();
                new CompositeSheet(docText, localizations).WriteToSpreadsheet(document, sheet);
            }
        }

        private string FullToLocalPath (string fullPath)
        {
            var prefix = fullPath.EndsWithFast(scriptFileExtension) ? scriptFolderPath : textFolderPath;
            return fullPath.Remove(prefix).TrimStart('\\').TrimStart('/');
        }

        private string LocalToFullPath (string localPath)
        {
            var prefix = localPath.EndsWithFast(scriptFileExtension) ? scriptFolderPath : textFolderPath;
            return $"{prefix}/{localPath}";
        }

        private string LocalPathToSheetName (string localPath)
        {
            var namePrefix = localPath.EndsWithFast(scriptFileExtension) ? scriptSheetNamePrefix : textSheetNamePrefix;
            return namePrefix + localPath.Replace("\\", sheetPathSeparator).Replace("/", sheetPathSeparator).GetBeforeLast(".");
        }

        private string SheetNameToLocalPath (string sheetName)
        {
            var namePrefix = sheetName.StartsWithFast(scriptSheetNamePrefix) ? scriptSheetNamePrefix : textSheetNamePrefix;
            var fileExtension = namePrefix == scriptSheetNamePrefix ? scriptFileExtension : textFileExtension;
            return sheetName.GetAfterFirst(namePrefix).Replace(sheetPathSeparator, "/") + fileExtension;
        }

        private Script LoadScriptAtPath (string scriptPath)
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

        private IReadOnlyCollection<string> LocateLocalizationsFor (string localPath, bool skipMissing = true)
        {
            if (!Directory.Exists(localeFolderPath)) return new string[0];

            var prefix = localPath.EndsWithFast(scriptFileExtension)
                ? ScriptsConfiguration.DefaultPathPrefix
                : ManagedTextConfiguration.DefaultPathPrefix;
            var paths = new List<string>();
            foreach (var localeDir in Directory.EnumerateDirectories(localeFolderPath))
            {
                var localizationPath = Path.Combine(localeDir, prefix, localPath);
                if (skipMissing && !File.Exists(localizationPath))
                {
                    Debug.LogWarning($"Missing localization resource for `{localPath}` (expected in `{localizationPath}`).");
                    continue;
                }

                paths.Add(localizationPath);
            }

            return paths;
        }
        
        private void NotifyProgressChanged (IReadOnlyList<string> paths, int index)
        {
            if (onProgress is null) return;
            
            var path = paths[index];
            var name = path.Contains(sheetPathSeparator) ? path : Path.GetFileNameWithoutExtension(path);
            var progress = index / (float)paths.Count;
            var info = $"Processing `{name}`...";
            var args = new ProgressChangedArgs(info, progress);
            onProgress.Invoke(args);
        }
    }
}
