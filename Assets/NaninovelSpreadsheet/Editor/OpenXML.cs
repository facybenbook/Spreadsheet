using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Naninovel.Spreadsheet
{
    // OpenXML documentation: https://docs.microsoft.com/en-us/office/open-xml/spreadsheets
    
    internal static class OpenXML
    {
        public static Worksheet GetOrAddSheet (this SpreadsheetDocument document, string sheetName)
        {
            var workbookPart = document.WorkbookPart;
            var sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName) ?? InsertSheet();
            return ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;

            Sheet InsertSheet ()
            {
                var newWorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
                newWorksheetPart.Worksheet = new Worksheet(new SheetData());
                var sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                var relationshipId = document.WorkbookPart.GetIdOfPart(newWorksheetPart);
                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Any())
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                var newSheet = new Sheet { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                sheets.AppendChild(newSheet);
                return newSheet;
            }
        }

        public static string GetCellValue (this SpreadsheetDocument document, Worksheet worksheet, string columnName, uint rowIndex)
        {
            var address = columnName + rowIndex;
            var cell = worksheet.GetFirstChild<SheetData>()
                .Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex)?
                .Elements<Cell>().FirstOrDefault(c => c.CellReference.Value.EqualsFast(address));
            if (cell?.DataType is null || cell.InnerText.Length == 0) return null;

            switch (cell.DataType.Value)
            {
                case CellValues.String:
                    return cell.CellValue.Text;
                case CellValues.SharedString:
                    var stringTable = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                    return stringTable?.SharedStringTable.ElementAt(int.Parse(cell.InnerText)).InnerText;
                case CellValues.Boolean:
                    return cell.InnerText.EqualsFast("0") ? "FALSE" : "TRUE";
                default:
                    return cell.InnerText;
            }
        }

        public static void SetCellValue (this SpreadsheetDocument document, Worksheet worksheet, string columnName, uint rowIndex, string value)
        {
            var sharedStringPart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() ??
                                   document.WorkbookPart.AddNewPart<SharedStringTablePart>();

            var index = InsertSharedStringItem(value, sharedStringPart);
            var cell = InsertCellInWorksheet();
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
            worksheet.Save();

            int InsertSharedStringItem (string text, SharedStringTablePart shareStringPart)
            {
                if (shareStringPart.SharedStringTable is null)
                    shareStringPart.SharedStringTable = new SharedStringTable();
                
                int itemIndex = 0;
                foreach (var item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
                {
                    if (item.InnerText == text) return itemIndex;
                    itemIndex++;
                }

                shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                shareStringPart.SharedStringTable.Save();
                return itemIndex;
            }

            Cell InsertCellInWorksheet ()
            {
                var sheetData = worksheet.GetFirstChild<SheetData>();
                var cellReference = columnName + rowIndex;

                var row = default(Row);
                if (sheetData.Elements<Row>().Count(r => r.RowIndex == rowIndex) != 0)
                    row = sheetData.Elements<Row>().First(r => r.RowIndex == rowIndex);
                else
                {
                    row = new Row { RowIndex = rowIndex };
                    sheetData.AppendChild(row);
                }

                if (row.Elements<Cell>().Any(c => c.CellReference.Value == columnName + rowIndex))
                    return row.Elements<Cell>().First(c => c.CellReference.Value == cellReference);

                var refCell = default(Cell);
                foreach (var c in row.Elements<Cell>())
                {
                    if (string.Compare(c.CellReference.Value, cellReference, StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        refCell = c;
                        break;
                    }
                }

                var newCell = new Cell { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);
                worksheet.Save();
                return newCell;
            }
        }
    }
    
}
