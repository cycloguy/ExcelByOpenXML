/*
//Reference from
//http://www.dispatchertimer.com/tutorial/how-to-create-an-excel-file-in-net-using-openxml-part-1-basics/
//
*/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelByOpenXML
{
    class Report
    {
        public void CreateExcelDoc(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument
                .Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                List<Employee> employees = Employees.EmployeesList;
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() {
                                            Id = workbookPart.GetIdOfPart(worksheetPart),
                                            SheetId = 1,
                                            Name = "Employees"
                                          };
                //Construct header
                Row row = new Row();
                row.Append(
                    ConstructCell("Id", CellValues.String),
                    ConstructCell("Name", CellValues.String),
                    ConstructCell("Birth Date", CellValues.String),
                    ConstructCell("Salary", CellValues.String));
                sheetData.AppendChild(row);
                //Add data
                foreach (var employee in employees)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(employee.Id.ToString(), CellValues.Number),
                        ConstructCell(employee.Name, CellValues.String),
                        ConstructCell(employee.DOB.ToString("yyyy/MM/dd"), CellValues.String),
                        ConstructCell(employee.Salary.ToString(CultureInfo.CurrentCulture), CellValues.Number));

                    sheetData.AppendChild(row);
                }
                sheets.AppendChild(sheet);
                //workbookPart.Workbook.Save();
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
