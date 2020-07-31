using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace excelExport
{
    class CustomSpreadsheet
    {
        public SpreadsheetDocument spreadsheet { get; set; }
        public WorkbookPart workbook { get; set; }
        public SharedStringTablePart sharedStrings { get; set; }
        
        public void Save()
        {
            this.spreadsheet.WorkbookPart.Workbook.Save();
            //this.workbook.
        }

        public void Close()
        {

            this.spreadsheet.Close();

            //catch (NullReferenceException ex)
            //{
            //    return;
            //}
            
        }

        //private WorkbookPart
        public CustomSpreadsheet(string path)
        {
            if (!System.IO.File.Exists(path))
            {
                this.spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
                try
                {
                    SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
                }
                catch(Exception ex)
                {
                    this.spreadsheet = null;
                    return;
                }
                

                WorkbookPart workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                this.workbook = workbookPart;

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);
                //this.Save();
            }
            else
            {
                OpenSettings openSettings = new OpenSettings();
                openSettings.MarkupCompatibilityProcessSettings =
                    new MarkupCompatibilityProcessSettings(
                        MarkupCompatibilityProcessMode.ProcessAllParts,
                        FileFormatVersions.Office2013
                    );
                try
                {
                    this.spreadsheet = SpreadsheetDocument.Open(path, true, openSettings);
                }
                catch(Exception ex)
                {
                    this.spreadsheet = null;
                }
                if(this.spreadsheet.GetPartsCountOfType<WorkbookPart>() > 0)
                    this.workbook = this.spreadsheet.WorkbookPart;
                else
                {
                    this.workbook = this.spreadsheet.AddWorkbookPart();
                }
                if(this.workbook == null)
                {
                    this.workbook = this.spreadsheet.AddWorkbookPart();
                    this.workbook.Workbook = new Workbook();
                    this.InsertWorksheet();
                }
                if (this.workbook.GetPartsCountOfType<SharedStringTablePart>() > 0)
                {
                    this.sharedStrings = this.workbook.GetPartsOfType<SharedStringTablePart>().First();
                }
            }
        }
        public Sheet AddNewSheet(string sheetName)
        {
            if (this.spreadsheet == null)
            {
                return null;
            }
            //Only need get first because a document shall have ony 1 workbook
            WorkbookPart workbookPart = this.spreadsheet.GetPartsOfType<WorkbookPart>().First();
            if (workbookPart == null)
            {
                workbookPart = spreadsheet.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
            }

            //  New sheet.xml file: which contains data of a sheet
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            //workbookPart.WorksheetParts.First();
            //  Add to workbook:
            Sheets sheets = spreadsheet.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            uint newSheetID;

            if(sheets == null)
            {
                sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                newSheetID = 1;
            }
            else
            {
                newSheetID = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1; 
                    //uint sID = uint.Parse(sheet.GetAttribute("sheetID", sheet.NamespaceUri).ToString());
            }
            Sheet newSheet;
            int isExist = sheets.Elements<Sheet>().Where(s => s.Name.ToString().Equals(sheetName)).Count();
            if(isExist == 0)
            {
                newSheet = new Sheet() { Id = this.spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = newSheetID, Name = sheetName };
                sheets.Append(newSheet);
            }
            else
            {
                return sheets.Elements<Sheet>().Where(s => s.Name.ToString().Equals(sheetName)).FirstOrDefault();
            }
            return newSheet;
        }

        public WorksheetPart GetWorksheetPartByName(string sheetName)
        {
            if(this.workbook == null || this.workbook.Workbook == null)
            {
                Console.WriteLine("Error: This spreadsheet has no workbook!");
                return null;
            }
            if(this.workbook.Workbook.Descendants<Sheets>().Count() == 0
                || this.workbook.Workbook.Descendants<Sheet>().Count() == 0)
            {
                Console.WriteLine("Error: This spreadsheet has no sheets!");
                return null;
            }

            Sheets sheets = this.workbook.Workbook.GetFirstChild<Sheets>();
            if(sheets == null)
            {
                Console.WriteLine("Error: this workbook has no sheets elements");
                return null;
            }
            //Sheet selectedSheet = sheets.Elements<Sheet>().First();
            //Console.WriteLine("Name: " + selectedSheet.Name);
            Sheet selectedSheet = null;
            if (sheets.Elements<Sheet>().Where(s => s.Name.ToString().Equals(sheetName)).Count() > 0)
                selectedSheet = sheets.Elements<Sheet>().Where(s => s.Name.ToString().Equals(sheetName)).FirstOrDefault();
            if (selectedSheet == null)
            {
                Console.WriteLine("Error: this workbook has no '{0}'",sheetName);
                return null;
            }

            string selectedSheetID = selectedSheet.Id;

            return (WorksheetPart)this.workbook.GetPartById(selectedSheetID);
        }

        // If Force == true: Get the first sheet to insert, regardless the name;
        public bool InsertText(string text, string columnName, uint rowIndex, WorksheetPart worksheetPart=null, bool force=false)
        {
            int index = this.InsertSharedStringItem(text);

            if(worksheetPart == null)
            {
                if(force == true)
                {   
                    if(this.workbook.GetPartsOfType<WorksheetPart>().Count() > 0)
                        worksheetPart = this.workbook.GetPartsOfType<WorksheetPart>().FirstOrDefault();
                }
                if( worksheetPart == null)
                {
                    Console.WriteLine("Error: This spreadsheet has no according worksheet!");
                    return false;
                }
            }
            Cell cell = InsertCellInWorksheet(columnName, rowIndex, worksheetPart);
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            worksheetPart.Worksheet.Save();
            return true;
        }

        public int InsertSharedStringItem (string text)
        {
            if(this.sharedStrings == null)
            {
                Console.WriteLine("Warning: This spreadsheet has no sharedStrings, auto created new one!");
                if(this.workbook.GetPartsCountOfType<SharedStringTablePart>() > 0)
                {
                    this.sharedStrings = this.workbook.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    this.sharedStrings = this.workbook.AddNewPart<SharedStringTablePart>();
                }
                //this.sharedStrings = this.spreadsheet.AddNewPart<SharedStringTablePart>();
            }
            if(this.sharedStrings.SharedStringTable == null)
            {
                this.sharedStrings.SharedStringTable = new SharedStringTable();
            }

            int i = 0;
            foreach (SharedStringItem item in this.sharedStrings.SharedStringTable.Elements<SharedStringItem>())
            {
                if(item.InnerText == text)
                {
                    return i;
                }
                i++;
            }
            this.sharedStrings.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            this.sharedStrings.SharedStringTable.Save();
            return i;
        }

        public WorksheetPart InsertWorksheet()
        {
            if(this.workbook == null)
            {
                Console.WriteLine("Error: this spreadsheet has no workbookPart!");
                return null;
            }
            WorksheetPart newWorksheetPart = this.workbook.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = this.workbook.Workbook.GetFirstChild<Sheets>();
            string relationshipID = this.workbook.GetIdOfPart(newWorksheetPart);

            uint sheetID = 1;
            if(sheets.Elements<Sheet>().Count() > 0)
            {
                sheetID = sheets.Elements<Sheet>().Select(s=> s.SheetId.Value).Max() +1;
            }
            string sheetName = "Sheet" + sheetID;

            Sheet sheet = new Sheet() { Id = relationshipID, SheetId = sheetID, Name = sheetName};
            sheets.Append(sheet);
            this.workbook.Workbook.Save();

            return newWorksheetPart;
        }

        public Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

    }
}
