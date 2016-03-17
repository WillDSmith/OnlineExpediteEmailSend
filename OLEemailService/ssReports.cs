using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data.Entity.SqlServer;

namespace OLEemailService
{
    class ssReports
    {
        InternalEntities db = new InternalEntities();
        
        public void CreateExcelDoc(string fileName)
        {
            DateTime dtt = DateTime.Now;
            dtt = dtt.AddHours(-12);
            
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document.
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart.
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = String.Format("{0:MMddyyyyhhmmss}", DateTime.Now)
                    
                };

                sheets.Append(sheet);

                workbookPart.Workbook.Save();

                var entries = db.OnlineExpedites.Where(d => d.CreationDate > dtt || d.DateSentTimeStamp == null);

                if (!entries.Any())
                    Environment.Exit(0);

                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Construct Header
                Row row = new Row();

                row.Append(
                    ConstructCell("Request Id", CellValues.String),
                    ConstructCell("EDP Tool Number", CellValues.String),
                    ConstructCell("Date", CellValues.String),
                    ConstructCell("Purchase Order", CellValues.String),
                    ConstructCell("Item", CellValues.String),
                    ConstructCell("Quantity", CellValues.String),
                    ConstructCell("Message", CellValues.String),
                    ConstructCell("By", CellValues.String),
                    ConstructCell("Response", CellValues.String),
                    ConstructCell("Germany Responder", CellValues.String));
                    

                // Insert the header row to the Sheet data
                sheetData.AppendChild(row);

                // Inserting each Entry
                foreach (var entry in entries)
                {
                    row = new Row();
                    row.Append(
                        ConstructCell(entry.RequestId.ToString(), CellValues.Number),
                        ConstructCell(entry.EDPToolNumber, CellValues.String),
                        ConstructCell(entry.CreationDate.ToString(), CellValues.String),
                        ConstructCell(entry.PurchaseOrderToGermany, CellValues.String),
                        ConstructCell(entry.LineNumber.ToString(), CellValues.Number),
                        ConstructCell(entry.QuantityRequested.ToString(), CellValues.Number),
                        ConstructCell(entry.Message, CellValues.String),
                        ConstructCell(entry.Name, CellValues.String),
                        ConstructCell(entry.Response, CellValues.String),
                        ConstructCell(entry.GermanyResponder, CellValues.String));
                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
                
            }
        }

        private Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }
        
    }
}
