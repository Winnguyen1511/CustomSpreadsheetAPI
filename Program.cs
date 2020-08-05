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
            

            //CustomSpreadsheet test_sheet = new CustomSpreadsheet("G:\\Csharp\\excelExport\\test_style.xlsx");
            //string sheetName = "mySheet2";
            //WorksheetPart newsheet = test_sheet.AddNewSheet(sheetName);
            //if (newsheet != null)
            //{
            //    Console.WriteLine("Successful add sheets");
            //}
            //else
            //{
            //    Console.WriteLine("Failed to add sheets");
            //    //return;
            //}
            //if (test_sheet.spreadsheet != null)
            //{
            //    test_sheet.Save();
            //    //test_sheet.Close();
            //}
            //WorksheetPart sheet1 = test_sheet.GetWorksheetPartByName(sheetName);
            //if (sheet1 == null)
            //{
            //    Console.WriteLine("There is no {0} in this speadsheet!", sheetName);
            //    //return;
            //}
            //else
            //{
            //    int maxNum = 10;
            //    for (int i = 1; i < maxNum; i++)
            //    {
            //        string numStr = "Num" + (i * 100).ToString();
            //        test_sheet.InsertText(numStr, "C", (uint)i, sheet1);
            //    }
            //    Console.WriteLine("Successful add {0} texts to {1}", maxNum, sheetName);
            //}
            //if (test_sheet.spreadsheet != null)
            //{
            //    test_sheet.Save();
            //    //test_sheet.Close();
            //}
            //Console.WriteLine(">> Test default sheet, Getting first 5 A-row:");
            //for (int i = 1; i <= 5; i++)
            //{
            //    string tmp = test_sheet.GetCellValue("A" + i, null, true);
            //    Console.WriteLine("A{0}: {1}", i, tmp);
            //}

            //Console.WriteLine(">> Test sheet, Getting first 15 C-row:");
            //WorksheetPart getSheet = test_sheet.GetWorksheetPartByName("mySheet2");
            //for (int i = 1; i <= 15; i++)
            //{
            //    string tmp = test_sheet.GetCellValue("C" + i, getSheet);
            //    Console.WriteLine("C{0}: {1}", i, tmp);
            //}

            //Console.WriteLine(">> Test sheet, Getting first 10 D-row:");
            //getSheet = test_sheet.GetWorksheetPartByName("mysheet");
            //for (int i = 1; i <= 15; i++)
            //{
            //    string tmp = test_sheet.GetCellValue("A" + i, getSheet);
            //    Console.WriteLine("D{0}: {1}", i, tmp);
            //}

            //Console.WriteLine(">> Test delete text, :");
            //getSheet = test_sheet.GetWorksheetPartByName("Sheet1");
            //for (int i = 1; i <= 4; i++)
            //{
            //    bool res = test_sheet.DeleteCell("A", (uint)i, getSheet);
            //    if (res == true)
            //        Console.WriteLine("Deleted cell at {0}!", "A" + i);
            //}

            //Console.WriteLine(">> Test insert number:");
            //for(int i = 1; i <= 5; i++)
            //{
            //    bool res = test_sheet.InsertValue((i * 100).ToString(), "B", (uint)i, getSheet, CellValues.Number);
            //    if(res == true)
            //    {
            //        Console.WriteLine("Insert {0} number at {1}!", i * 100, "B" + i);
            //    }
            //}
            CustomSpreadsheet test_sheet = new CustomSpreadsheet("G:\\Csharp\\excelExport\\test_formula.xlsx");
            string sheetName = "Sheet1";
            WorksheetPart Sheet1 = test_sheet.GetWorksheetPartByName(sheetName);

            Cell refCell =test_sheet.InsertFormula("SUM(A1,B1)/7+A1", "D", 1, Sheet1);

            if(refCell != null)
            {
                Console.WriteLine("Successful added formula to C1!");
            }
            Console.WriteLine("Tess add Calculation Chain!");
            bool res = test_sheet.InsertFormulaChain("SUM(A2,B2)/7+A2", "D2", "D5",Sheet1);
            if(res == true)
            {
                Console.WriteLine("Successful added chain of calculation!");
            }
            if (test_sheet.spreadsheet != null)
            {
                test_sheet.Save();
                test_sheet.Close();
            }

            Console.ReadKey();
        }
    }
}
