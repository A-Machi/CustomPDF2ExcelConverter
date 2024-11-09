using CustomPDF2ExcelConverter.Model;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CustomPDF2ExcelConverter.Controller
{
    public class MoveDeletedDataIntoAnotherWorksheetInExcel
    {
        public static void MoveToAnotherWorksheet(IReadOnlyList<RetrievalDataDto> toBeMovedData, SheetData sourceSheetData, SheetData destinationSheetData)
        {
            foreach (var moveData in toBeMovedData)
            {
                var sourceRow = sourceSheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex! == moveData.RowNumber);

                if (sourceRow != null)
                {
                    var newRow = new Row() { RowIndex = (destinationSheetData.Elements<Row>().Select(r => r.RowIndex!.Value).Max() + 1) };
                    foreach (var cell in sourceRow.Elements<Cell>())
                    {
                        var newCell = cell.CloneNode(true) as Cell;
                        if (newCell is null || newCell.CellReference is null || newCell.CellReference.Value is null)
                        {
                            throw new Exception("New Cell is null");
                        }
                        newCell.CellReference = newCell.CellReference.Value.Replace(moveData.RowNumber.ToString(), newRow.RowIndex.ToString());
                        newRow.Append(newCell);
                    }
                    destinationSheetData.Append(newRow);
                }
                else
                {
                    throw new ArgumentOutOfRangeException(nameof(moveData.RowNumber), "The specified row does not exist in the source worksheet.");
                }
            }
        }

        public static void DeleteFromSourceWorksheet(SheetData sheetData, IReadOnlyList<RetrievalDataDto> toBeMovedData)
        {
            var orderedDeletedData = toBeMovedData.OrderBy(x => x.RowNumber).GroupBy(x => x.OrderNumber);

            foreach (var deletedData in orderedDeletedData)
            {
                foreach (var data in deletedData)
                {
                    var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex! == data.RowNumber);
                    if (row != null)
                    {
                        row.Remove();
                    }
                }  
            }
        }
    }
}
