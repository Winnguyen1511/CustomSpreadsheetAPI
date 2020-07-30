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
        public void Save()
        {
            this.spreadsheet.WorkbookPart.Workbook.Save();
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

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = spreadsheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
                sheets.Append(sheet);
                this.Save();
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
                return null;
            }
            return newSheet;
        }
    }
}
