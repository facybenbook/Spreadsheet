using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using UnityEngine;

namespace Naninovel.Spreadsheet
{
    public class SpreadsheetProcessor
    {
        public class Parameters
        {
            public string SpreadsheetPath { get; set; }
            public bool SingleSpreadsheet { get; set; }
            public string ScriptFolderPath { get; set; }
            public string ManagedTextFolderPath { get; set; }
            public string LocalizationFolderPath { get; set; }
        }

        private const string scriptsCategory = "Scripts";
        private const string textCategory = "Text";
        private const string sheetPathSeparator = ">";
        private const string scriptSheetNamePrefix = scriptsCategory + sheetPathSeparator;
        private const string textSheetNamePrefix = textCategory + sheetPathSeparator;
        private const string scriptFileExtension = ".nani";
        private const string textFileExtension = ".txt";
        private const string scriptFilePattern = "*" + scriptFileExtension;
        private const string textFilePattern = "*" + textFileExtension;

        private readonly string spreadsheetPath;
        private readonly bool singleSpreadsheet;
        private readonly string scriptFolderPath;
        private readonly string textFolderPath;
        private readonly string localeFolderPath;
        private readonly Action<ProgressChangedArgs> onProgress;

        public SpreadsheetProcessor (Parameters parameters, Action<ProgressChangedArgs> onProgress = default)
        {
            spreadsheetPath = parameters.SpreadsheetPath;
            singleSpreadsheet = parameters.SingleSpreadsheet;
            scriptFolderPath = parameters.ScriptFolderPath;
            textFolderPath = parameters.ManagedTextFolderPath;
            localeFolderPath = parameters.LocalizationFolderPath;
            this.onProgress = onProgress;
        }

        public void Export ()
        {
            if (Directory.Exists(scriptFolderPath))
                ExportScripts();
            if (Directory.Exists(textFolderPath))
                ExportManagedText();
        }

        public void Import ()
        {
            var directory = Path.GetDirectoryName(spreadsheetPath);
            var documentPaths = Directory.GetFiles(directory, "*.xlsx", SearchOption.AllDirectories);
            for (int i = 0; i < documentPaths.Length; i++)
            {
                var document = SpreadsheetDocument.Open(documentPaths[i], false);
                var sheetsNames = document.GetSheetNames().ToArray();
                for (int j = 0; j < sheetsNames.Length; j++)
                {
                    if (singleSpreadsheet) NotifyProgressChanged(sheetsNames, j);
                    else NotifyProgressChanged(documentPaths, i);
                    ImportSheet(document, sheetsNames[j], documentPaths[i]);
                }
                document.Dispose();
            }
        }

        private void ImportSheet (SpreadsheetDocument document, string sheetName, string docPath)
        {
            var sheet = document.GetSheet(sheetName);
            var localPath = SheetNameToLocalPath(sheetName);
            if (localPath is null)
            {
                Debug.LogWarning($"Sheet `{sheetName}` in `{docPath}` is not recognized and will be ignored.");
                return;
            }
            var fullPath = LocalToFullPath(localPath);
            var localizations = LocateLocalizationsFor(localPath, false);
            var compositeSheet = new CompositeSheet(document, sheet);
            var managedText = fullPath.EndsWithFast(textFileExtension);
            compositeSheet.WriteToProject(fullPath, localizations, managedText);
        }

        private SpreadsheetDocument OpenOrCreateDocument (string category, string localPath)
        {
            if (singleSpreadsheet) return SpreadsheetDocument.Open(spreadsheetPath, true);
            var directory = Path.Combine(spreadsheetPath, category, Path.GetDirectoryName(localPath));
            if (!Directory.Exists(directory)) Directory.CreateDirectory(directory);
            var path = Path.Combine(directory, Path.GetFileNameWithoutExtension(localPath) + ".xlsx");
            return OpenXML.CreateDocument(path);
        }

        private void ExportScripts ()
        {
            var scriptPaths = Directory.GetFiles(scriptFolderPath, scriptFilePattern, SearchOption.AllDirectories);
            for (int pathIndex = 0; pathIndex < scriptPaths.Length; pathIndex++)
            {
                NotifyProgressChanged(scriptPaths, pathIndex);
                var scriptPath = scriptPaths[pathIndex];
                var script = LoadScriptAtPath(scriptPath);
                var localPath = FullToLocalPath(scriptPath);
                var document = OpenOrCreateDocument(scriptsCategory, localPath);
                var sheetName = LocalPathToSheetName(localPath);
                var sheet = document.GetSheet(sheetName) ?? document.AddSheet(sheetName);
                var localizations = LocateLocalizationsFor(localPath).Select(LoadScriptAtPath).ToArray();
                new CompositeSheet(script, localizations).WriteToSpreadsheet(document, sheet);
                document.Dispose();
            }
        }

        private void ExportManagedText ()
        {
            var managedTextPaths = Directory.GetFiles(textFolderPath, textFilePattern, SearchOption.AllDirectories);
            for (int pathIndex = 0; pathIndex < managedTextPaths.Length; pathIndex++)
            {
                NotifyProgressChanged(managedTextPaths, pathIndex);
                var docPath = managedTextPaths[pathIndex];
                var docText = File.ReadAllText(docPath);
                var localPath = FullToLocalPath(docPath);
                var document = OpenOrCreateDocument(textCategory, localPath);
                var sheetName = LocalPathToSheetName(localPath);
                var sheet = document.GetSheet(sheetName) ?? document.AddSheet(sheetName);
                var localizations = LocateLocalizationsFor(localPath).Select(File.ReadAllText).ToArray();
                new CompositeSheet(docText, localizations).WriteToSpreadsheet(document, sheet);
                document.Dispose();
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
            if (!sheetName.Contains(namePrefix)) return null;
            var fileExtension = namePrefix == scriptSheetNamePrefix ? scriptFileExtension : textFileExtension;
            return sheetName.GetAfterFirst(namePrefix).Replace(sheetPathSeparator, "/") + fileExtension;
        }

        private ScriptText LoadScriptAtPath (string scriptPath)
        {
            var scriptText = File.ReadAllText(scriptPath);
            var script = Script.FromScriptText(scriptPath, scriptText);
            if (script == null) throw new Exception($"Failed to load `{scriptPath}` script.");
            return new ScriptText(script, scriptText);
        }

        private IReadOnlyCollection<string> LocateLocalizationsFor (string localPath, bool skipMissing = true)
        {
            if (!Directory.Exists(localeFolderPath)) return Array.Empty<string>();

            var prefix = localPath.EndsWithFast(scriptFileExtension)
                ? ScriptsConfiguration.DefaultPathPrefix
                : ManagedTextConfiguration.DefaultPathPrefix;
            var paths = new List<string>();
            foreach (var localeDir in Directory.EnumerateDirectories(localeFolderPath))
            {
                var localizationPath = Path.Combine(localeDir, prefix, localPath).Replace('\\', '/');
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
