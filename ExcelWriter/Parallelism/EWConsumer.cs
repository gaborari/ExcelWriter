using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelWriter.Entities;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelWriter.Parallelism
{
    internal class EWConsumer
    {
        private ConcurrentQueue<EWCell> _cellQueue;
        private object _lockObj;
        private string fileName;
        private EWThreadKiller threadKiller;
        private EWProcessingFinished finished;
        private List<EWSheet> _sheet;

        int sheetIndex;
        int rowIndex = 1;
        int colIndex;


        public EWConsumer(List<EWSheet> sheets, ConcurrentQueue<EWCell> _cellQueue, object _lockObj, string fileName, EWThreadKiller threadKiller, EWProcessingFinished finished)
        {
            this._cellQueue = _cellQueue;
            this._lockObj = _lockObj;
            this.fileName = fileName;
            this.threadKiller = threadKiller;
            this.finished = finished;
            this._sheet = sheets;
        }

        public void consume()
        {
            EWCell item;
            
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName), SpreadsheetDocumentType.Workbook))
            {
                spreadSheet.AddWorkbookPart();
                spreadSheet.WorkbookPart.Workbook = new Workbook();     // create the worksheet
                // create the worksheet to workbook relation
                spreadSheet.WorkbookPart.Workbook.AppendChild(new Sheets());

               
                //Init base style
                WorkbookStylesPart stylesPart = spreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = EWStyle.GetBaseStyle();
                stylesPart.Stylesheet.Save();

                foreach (var currSheet in _sheet)
                {

                    #region init stuff
                    WorkbookPart workbookPart = spreadSheet.WorkbookPart;
                    ////

                    // Add a blank WorksheetPart.
                    WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    newWorksheetPart.Worksheet = new Worksheet(new SheetData());

                    Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
                    string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

                    // Get a unique ID for the new worksheet.
                    uint sheetId = 1;
                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                    }

                    // Give the new worksheet a name.
                    string sheetName = currSheet.Name;

                    // Append the new worksheet and associate it with the workbook.
                    Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                    sheets.Append(sheet);

                    newWorksheetPart.Worksheet.Save();

                    //
                    ////////


                    string origninalSheetId = workbookPart.GetIdOfPart(newWorksheetPart);


                    WorksheetPart replacementPart = workbookPart.AddNewPart<WorksheetPart>();
                    string replacementPartId = workbookPart.GetIdOfPart(replacementPart);

                    OpenXmlReader reader = OpenXmlReader.Create(newWorksheetPart);
                    OpenXmlWriter writer = OpenXmlWriter.Create(replacementPart);
                    
                    Row row = new Row();
                    Cell cell = new Cell() { CellValue = new CellValue("hellowrold") };

                    Worksheet worksheet = newWorksheetPart.Worksheet; 
                    #endregion
                    while (reader.Read())
                    {
                        if (reader.ElementType == typeof(SheetData))
                        {
                            if (reader.IsEndElement)
                            {
                                //writer.WriteMergedCells(currSheet._mergedCells);
                                
                                //write merged cells

                                continue;
                            }
                            writer.WriteStartElement(new SheetData());
                            //write first row
                            writer.WriteStartElement(row);

                            while (true)
                            {
                                lock (_lockObj)
                                {
                                    if (!_cellQueue.IsEmpty)
                                    {
                                        _cellQueue.TryDequeue(out item);
                                        
                                        //Row and col process logic

                                        //staying in the current row
                                        if (item.rowIndex == rowIndex)
                                        {
                                            //write the cell
                                            cell.CellValue.Text = item.value;
                                            cell.StyleIndex = 0;
                                            writer.WriteElement(cell);

                                        }
                                        //create new row
                                        else
                                        {
                                            //Close the previous row
                                            writer.WriteEndElement();

                                            //write new row
                                            writer.WriteStartElement(row);

                                            //set to the actual row
                                            rowIndex = item.rowIndex;
                                        }
                                    }
                                    else
                                    {
                                        if (this.threadKiller.Killed == true)
                                        {
                                            break;
                                        }

                                        Monitor.Wait(_lockObj);
                                        continue;
                                    }
                                }
                            }
                            //close last row
                            writer.WriteEndElement();
                            //close sheet
                            writer.WriteEndElement();
                        }
                        #region closing stuff
                        else
                        {
                            if (reader.IsStartElement)
                            {
                                writer.WriteStartElement(reader);
                            }
                            else if (reader.IsEndElement)
                            {
                                writer.WriteEndElement();
                            }
                        }
                    }

                    reader.Close();
                    writer.Close();

                    Sheet sheeet = workbookPart.Workbook.Descendants<Sheet>().Where(s => s.Id.Value.Equals(origninalSheetId)).First();
                    sheeet.Id.Value = replacementPartId;
                    workbookPart.DeletePart(newWorksheetPart); 
                        #endregion
                }

                lock (_lockObj)
                {
                    finished.EWFinished = true;
                }

                //
                spreadSheet.WorkbookPart.Workbook.Save();
            }

            
        }
    }
}
