using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelWriter.Parallelism
{
    internal class EWConsumer
    {
        private string _fileName { get; set; }
       
        public EWConsumer(string fileName)
        {
            _fileName = fileName;
        }

        public void consume()
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, _fileName), SpreadsheetDocumentType.Workbook))
            {
                List<OpenXmlAttribute> attributeList = new List<OpenXmlAttribute>();
                OpenXmlWriter writer;

                spreadSheet.AddWorkbookPart();
                //Init base style
                WorkbookStylesPart stylesPart = spreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = ExcelWriter.GetStyleSheet();
                stylesPart.Stylesheet.Save();

                var currentSheetIndex = 0;
                while (currentSheetIndex < ExcelWriter.Sheets.Count())
                {
                    var sheet = ExcelWriter.Sheets.ElementAt(currentSheetIndex);

                    WorksheetPart worksheetPart = spreadSheet.WorkbookPart.AddNewPart<WorksheetPart>();

                    writer = OpenXmlWriter.Create(worksheetPart);
                    writer.WriteStartElement(new Worksheet());
                    {
                        writer.WriteStartElement(new SheetData());

                        int rowIndex = 0;
                        writer.WriteStartElement(new Row(), attributeList);
                        EWCell item;

                        while (ExcelWriter.Finished == false)
                        {
                            if (!ExcelWriter.CellQueue.IsEmpty)
                            {
                                ExcelWriter.CellQueue.TryDequeue(out item);
                                ExcelWriter.RemoveCellCount++;

                                if (item.sheetIndex != currentSheetIndex)
                                {
                                    break;
                                }

                                if (item.rowIndex == rowIndex)
                                {
                                    attributeList = new List<OpenXmlAttribute>();
                                    // this is the data type ("t"), with CellValues.String ("str")
                                    attributeList.Add(new OpenXmlAttribute("t", null, "str"));
                                    //attributeList.Add(new OpenXmlAttribute("s", null, (UInt32Value)1U));
                                    attributeList.Add(new OpenXmlAttribute("s", null, "1"));

                                    writer.WriteStartElement(new Cell(), attributeList);

                                    writer.WriteElement(new CellValue(item.value));

                                    // this is for Cell
                                    writer.WriteEndElement();
                                }
                                else
                                {
                                    // this is for Row
                                    writer.WriteEndElement();
                                    writer.WriteStartElement(new Row(), attributeList);
                                    rowIndex = item.rowIndex;
                                }

                            }
                            else
                            {
                                if (ExcelWriter.AddingCellsInProgress == false)
                                {
                                    ExcelWriter.Finished = true;
                                    writer.WriteEndElement();
                                    GC.Collect(); 
                                }
                                //TODO Write merged cells
                            }
                        }

                        // this is for SheetData
                        writer.WriteEndElement();
                    }
                    // this is for Worksheet
                    writer.WriteEndElement();
                    writer.Close();
                    currentSheetIndex++;
                }

                writer = OpenXmlWriter.Create(spreadSheet.WorkbookPart);
                writer.WriteStartElement(new Workbook());
                {
                    writer.WriteStartElement(new Sheets());

                    currentSheetIndex = 0;
                    foreach (var sheet in ExcelWriter.Sheets)
                    {
                        writer.WriteElement(new Sheet()
                        {
                            Name = sheet.Name,
                            SheetId = Convert.ToUInt32(currentSheetIndex + 1),
                            Id = spreadSheet.WorkbookPart.GetIdOfPart(spreadSheet.WorkbookPart.WorksheetParts.ElementAt(currentSheetIndex))
                        });
                        currentSheetIndex++;
                    }

                    // this is for Sheets
                    writer.WriteEndElement();
                }
                // this is for Workbook
                writer.WriteEndElement();

                writer.Close();

                spreadSheet.Close();
            }


        }
    }
}
