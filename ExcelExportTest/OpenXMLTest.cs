using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Util;
using DataTable = System.Data.DataTable;

namespace ExportUsingOpenXML
{
    class OpenXMLTest
    {
        public void DoTest()
        {
            var gdt = new GenerateDataTable();
            var table = gdt.GetNewTable();

            Stopwatch sw = new Stopwatch();
            sw.Start();
            var filename = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx");
            WriteExcelFile(filename, table);
            sw.Stop();
            Console.WriteLine($"Save used {sw.ElapsedMilliseconds} ms");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private static void WriteExcelFile(string filename, DataTable table)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = table.TableName };

                sheets.Append(sheet);

                Row headerRow = new Row();

                List<String> columns = new List<string>();
                foreach (System.Data.DataColumn column in table.Columns)
                {
                    columns.Add(column.ColumnName);

                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(column.ColumnName);
                    headerRow.AppendChild(cell);
                }

                sheetData.AppendChild(headerRow);

                foreach (DataRow dsrow in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (String col in columns)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.String;
                        cell.CellValue = new CellValue(dsrow[col].ToString());
                        newRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(newRow);
                }

                workbookPart.Workbook.Save();
            }
        }


        public void DoTest2()
        {
            var gdt = new GenerateDataTable();
            var table = gdt.GetNewTable(1000000);

            Stopwatch sw = new Stopwatch();
            sw.Start();
            var filename = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xlsx");
            LargeExport(filename, table);
            sw.Stop();
            Console.WriteLine($"Save used {sw.ElapsedMilliseconds} ms");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        public static void LargeExport(string filename, DataTable table)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook))
            {
                //this list of attributes will be used when writing a start element
                List<OpenXmlAttribute> attributes;
                OpenXmlWriter writer;

                document.AddWorkbookPart();
                WorksheetPart workSheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();

                writer = OpenXmlWriter.Create(workSheetPart);
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());

                // 表头列信息
                //create a new list of attributes
                attributes = new List<OpenXmlAttribute>();
                // add the row index attribute to the list
                attributes.Add(new OpenXmlAttribute("r", null, "1"));
                //write the row start element with the row index attribute
                writer.WriteStartElement(new Row(), attributes);
                for (int columnNum = 0; columnNum < table.Columns.Count; ++columnNum)
                {
                    //reset the list of attributes
                    attributes = new List<OpenXmlAttribute>();
                    // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                    attributes.Add(new OpenXmlAttribute("t", null, "str"));
                    //add the cell reference attribute
                    attributes.Add(new OpenXmlAttribute("r", "", $"{GetColumnName(columnNum+1)}1"));

                    //write the cell start element with the type and reference attributes
                    writer.WriteStartElement(new Cell(), attributes);
                    //write the cell value
                    writer.WriteElement(new CellValue(table.Columns[columnNum].ColumnName));

                    // write the end cell element
                    writer.WriteEndElement();
                }
                // write the end row element
                writer.WriteEndElement();

                for (int rowNum = 1; rowNum <= table.Rows.Count; ++rowNum)
                {
                    int docRowNum = rowNum + 1;
                    //create a new list of attributes
                    attributes = new List<OpenXmlAttribute>();
                    // add the row index attribute to the list
                    attributes.Add(new OpenXmlAttribute("r", null, docRowNum.ToString()));

                    //write the row start element with the row index attribute
                    writer.WriteStartElement(new Row(), attributes);

                    for (int columnNum = 1; columnNum <= table.Columns.Count; ++columnNum)
                    {
                        //reset the list of attributes
                        attributes = new List<OpenXmlAttribute>();
                        // add data type attribute - in this case inline string (you might want to look at the shared strings table)
                        attributes.Add(new OpenXmlAttribute("t", null, "str"));
                        //add the cell reference attribute
                        attributes.Add(new OpenXmlAttribute("r", "", $"{GetColumnName(columnNum)}{docRowNum}"));

                        //write the cell start element with the type and reference attributes
                        writer.WriteStartElement(new Cell(), attributes);
                        var cellValue = table.Rows[rowNum - 1][columnNum - 1];
                        string cellStr = cellValue == null ? "" : (cellValue is DateTime?((DateTime)cellValue).ToString("yyyy-MM-dd HH:mm:ss.fff") : cellValue.ToString());
                        //write the cell value
                        writer.WriteElement(new CellValue(cellStr));

                        // write the end cell element
                        writer.WriteEndElement();
                    }

                    // write the end row element
                    writer.WriteEndElement();
                }

                // write the end SheetData element
                writer.WriteEndElement();
                // write the end Worksheet element
                writer.WriteEndElement();
                writer.Close();

                writer = OpenXmlWriter.Create(document.WorkbookPart);
                writer.WriteStartElement(new Workbook());
                writer.WriteStartElement(new Sheets());

                writer.WriteElement(new Sheet()
                {
                    Name = table.TableName,
                    SheetId = 1,
                    Id = document.WorkbookPart.GetIdOfPart(workSheetPart)
                });

                // End Sheets
                writer.WriteEndElement();
                // End Workbook
                writer.WriteEndElement();

                writer.Close();

                document.Close();
            }
        }

        //A simple helper to get the column name from the column index. This is not well tested!
        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;
            int modifier;

            while (dividend > 0)
            {
                modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }
    }
}
