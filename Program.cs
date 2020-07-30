using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace excelExport
{
    class Program
    {   
        public static void CreateSpreadsheetWorkBook(string filepath)
        {
            OpenSettings openSettings = new OpenSettings();
            openSettings.MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    FileFormatVersions.Office2013
                );

            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(filepath,SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);

            //  Adding

            workbookPart.Workbook.Save();
            spreadsheet.Close();
        }
        public static void CreateAddDataSpreadsheetWorkBook(string filepath)
        {
            OpenSettings openSettings = new OpenSettings();
            openSettings.MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    FileFormatVersions.Office2013
                );

            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

            WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
            sheets.Append(sheet);

            //  Adding data:
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            //  Add a row:
            Row row;
            row = new Row() { RowIndex = 1 };
            sheetData.Append(row);

            //  Find the A1 cell in the first Cell
            Cell refCell = null; 
            foreach(Cell cell in row.Elements<Cell>())
            {
                if(string.Compare(cell.CellReference.Value, "A1", true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }
            Cell newCell = new Cell() { CellReference = "A1" };
            row.InsertBefore(newCell, refCell);

            //  Set the cell value:
            newCell.CellValue = new CellValue("100");
            //Can be Shared String here:
            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);

            workbookPart.Workbook.Save();
            spreadsheet.Close();
        }
        public static void AddDataSpreadsheetWorkBook(string filepath, string col, uint rowNum, string text, CellValues type=CellValues.Number)
        {
            OpenSettings openSettings = new OpenSettings();
            openSettings.MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    FileFormatVersions.Office2013
                );

            SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filepath,true, openSettings);

            WorkbookPart workbookPart = spreadsheet.GetPartsOfType<WorkbookPart>().First();
            //workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.GetPartsOfType<WorksheetPart>().First();
            //worksheetPart.Worksheet = new Worksheet(new SheetData());

            //Sheets sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            //Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };

            //sheets.Append(sheet);

            //  Adding data:
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            //  Add a row:
            Row row;
            //row = new Row() { RowIndex = 1 };
            //sheetData.Append(row);
            try
            {
                row = (Row)sheetData.ElementAt((int)rowNum);
            }
            catch(ArgumentOutOfRangeException ex)
            {
                row = new Row() { RowIndex = new UInt32Value((uint)rowNum) };
                sheetData.Append(row);
            }
            //Row Row2nd = (Row)sheetData.ElementAt(100);
            
            //  Find the A1 cell in the first Cell
            Cell refCell = null;
            string cellLocation = col + rowNum.ToString();
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellLocation, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }
            Cell newCell = new Cell() { CellReference = cellLocation };
            row.InsertBefore(newCell, refCell);
            //row.InsertBefore()
            //  Set the cell value:
            newCell.CellValue = new CellValue(text);
            //Can be Shared String here:
            
            newCell.DataType = new EnumValue<CellValues>(type);

            workbookPart.Workbook.Save();
            spreadsheet.Close();
        }

        static void Main(string[] args)
        {
            //CreateSpreadsheetWorkBook("G:\\Csharp\\excelExport\\test.xlsx");
            //CreateAddDataSpreadsheetWorkBook("G:\\Csharp\\excelExport\\test_sheets.xlsx");
            //for(uint i = 2; i < 10; i++)
            //{
            //    AddDataSpreadsheetWorkBook("G:\\Csharp\\excelExport\\test_sheets.xlsx", "A", i, "200");
            //}

            CustomSpreadsheet test_sheet = new CustomSpreadsheet("G:\\Csharp\\excelExport\\test_sheets.xlsx");
            test_sheet.AddNewSheet("mynewsheet");
            test_sheet.Save();
            test_sheet.Close();

            Console.ReadKey();
        }
    }
}
