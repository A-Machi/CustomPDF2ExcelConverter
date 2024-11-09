using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;

namespace CustomPDF2ExcelConverter.Controller
{
    public class ConvertPDF2Excel
    {
        public static bool CustomPDF2ExcelConverterHandler(string pathToPDF, string pathToExcel)
        {
            var filePathPDF = new FileInfo(pathToPDF);
            var filePathToExcel = new FileInfo(pathToExcel);

            if (!filePathPDF.Exists)
            {
                throw new Exception("Did not find the PDF to read from.");

            }
            if (!filePathToExcel.Exists)
            {
                throw new Exception("Did not find the Excel to insert data into.");
            }
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(pathToExcel, true))
                {
                    var workbookPart = document.WorkbookPart;

                    if (workbookPart is null)
                    {
                        throw new Exception("WorkbookPart is null");
                    }

                    var retrievalWorksheet = GetWorksheet(workbookPart, "läuft gerade");
                    var movedRetrievalWorksheet = GetWorksheet(workbookPart, "kein Abruf");

                    var retrievalWorksheetData = retrievalWorksheet.GetFirstChild<SheetData>();
                    var movedRetrievalWorksheetData = movedRetrievalWorksheet.GetFirstChild<SheetData>();

                    if (retrievalWorksheetData is null)
                    {
                        throw new Exception("läuft gerade");
                    }
                    if (movedRetrievalWorksheetData is null)
                    {
                        throw new Exception("kein Abruf");
                    }

                    var extractedTextFromPDF = ReadDataFromPDF.ExtractTextFromPDF(pathToPDF);
                    var currentRetrievalDataStateInExcel = ReadDataFromExcel.GetExcelData(retrievalWorksheetData, workbookPart);

                    var (oldData, toReplaceOldData, toBeCreatedData, toBeMovedData) = ComparisonCheck.CompareData(extractedTextFromPDF, currentRetrievalDataStateInExcel);

                    if (toBeMovedData.Count > 0)
                    {
                        MoveDeletedDataIntoAnotherWorksheetInExcel.MoveToAnotherWorksheet(toBeMovedData, retrievalWorksheetData, movedRetrievalWorksheetData);
                        MoveDeletedDataIntoAnotherWorksheetInExcel.DeleteFromSourceWorksheet(retrievalWorksheetData, toBeMovedData);
                    }

                    if (toReplaceOldData.Count > 0 && oldData.Count > 0)
                    {
                        UpdateDataInExcel.UpdateDataFromWorksheet(retrievalWorksheetData, oldData, toReplaceOldData);
                    }

                    if (toBeCreatedData.Count > 0)
                    {
                        var currentMovedRetrievalDataStateInExcel = ReadDataFromExcel.GetExcelData(movedRetrievalWorksheetData, workbookPart);
                        CreateDataInExcel.InsertIntoExcelFile(retrievalWorksheetData, toBeCreatedData, currentMovedRetrievalDataStateInExcel);
                    }

                    ResetRowNumbers(retrievalWorksheetData);

                    retrievalWorksheet.Save();
                    movedRetrievalWorksheet.Save();
                    return true;
                }  
            }
            catch (IOException)
            {
                throw new Exception("Excel file is open. Please, close it!");
            }
            catch (Exception ex)
            {
                throw new Exception("An error occurred: " + ex.Message);
            }
        }

        private static Worksheet GetWorksheet(WorkbookPart workbookPart, string sheetName)
        {
            var sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            var sheet = sheets!.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);

            if (sheet != null)
            {
                var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
                return worksheetPart.Worksheet;
            }
            else
            {
                throw new Exception($"Worksheet with name '{sheetName}' does not exist.");
            }
        }

        private static void ResetRowNumbers(SheetData sheetData)
        {
            var currentRowIndex = 1;
            foreach (var row in sheetData.Elements<Row>())
            {
                row.RowIndex = (uint)currentRowIndex;

                var currentCellIndex = 1;
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellColumn = GetColumnName(cell.CellReference!);
                    cell.CellReference = $"{cellColumn}{currentRowIndex}";
                    currentCellIndex++;
                }

                currentRowIndex++;
            }
        }

        private static string GetColumnName(string cellReference)
        {
            int i;
            for (i = 0; i < cellReference.Length && !char.IsDigit(cellReference[i]); i++) ;
            return cellReference.Substring(0, i);
        }

        public static void StartExcel(string filePath)
        {
            try { 
                var startInfo = new ProcessStartInfo
                {
                    FileName = "excel.exe",
                    Arguments = filePath,
                    UseShellExecute = true
                };

                Process.Start(startInfo);
            } 
            catch(Exception er) 
            {
                throw new Exception("Could not start the program excel. Error: " + er.Message);
            }
        }
    }
}
