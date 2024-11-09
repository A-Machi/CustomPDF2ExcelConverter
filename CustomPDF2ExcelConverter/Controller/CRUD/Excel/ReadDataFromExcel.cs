using CustomPDF2ExcelConverter.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace CustomPDF2ExcelConverter.Controller
{
    public class ReadDataFromExcel
    {
        public static IReadOnlyList<RetrievalDataDto> GetExcelData(SheetData sheetData, WorkbookPart workbookPart)
        {
            var retrievalDataDto = new List<RetrievalDataDto>();
            var rows = sheetData.Descendants<Row>().Skip(1).ToList();

            foreach (var row in rows)
            {
                var cells = new List<Cell>(new Cell[26]);
                foreach (var cell in row.Elements<Cell>())
                {
                    var cellReference = cell.CellReference;
                    int cellIndex = GetColumnIndexFromCellReference(cellReference);
                    if (cellIndex >= 0 && cellIndex < 26)
                    {
                        cells[cellIndex] = cell;
                    }
                }

                var retrievalData = new RetrievalDataDto
                {
                    RowNumber = row.RowIndex!.Value,
                    OrderNumber = GetCellValue(cells.ElementAtOrDefault(0), workbookPart),
                    PTSOrder = GetCellValue(cells.ElementAtOrDefault(1), workbookPart),
                    Naming = GetCellValue(cells.ElementAtOrDefault(2), workbookPart),
                    Plant = GetCellValue(cells.ElementAtOrDefault(3), workbookPart),
                    UnloadingPoint = GetCellValue(cells.ElementAtOrDefault(4), workbookPart),
                    ItemNumberCustomer = GetCellValue(cells.ElementAtOrDefault(5), workbookPart),
                    Status = GetCellValue(cells.ElementAtOrDefault(6), workbookPart),
                    LastDelivery = GetCellValue(cells.ElementAtOrDefault(7), workbookPart),
                    WECaptureDate = GetCellValue(cells.ElementAtOrDefault(8), workbookPart),
                    VNumber = GetCellValue(cells.ElementAtOrDefault(9), workbookPart),
                    Appointment = GetCellValue(cells.ElementAtOrDefault(10), workbookPart),
                    Quantity = GetCellValue(cells.ElementAtOrDefault(11), workbookPart),
                    QuantityChange = GetCellValue(cells.ElementAtOrDefault(12), workbookPart),
                    Changes = GetCellValue(cells.ElementAtOrDefault(13), workbookPart),
                    Link = GetCellValue(cells.ElementAtOrDefault(14), workbookPart),
                    Remark = GetCellValue(cells.ElementAtOrDefault(15), workbookPart),
                    PriceBetween1_4 = GetCellValue(cells.ElementAtOrDefault(16), workbookPart),
                    PriceBetween5_9 = GetCellValue(cells.ElementAtOrDefault(17), workbookPart),
                    PriceBetween10_24 = GetCellValue(cells.ElementAtOrDefault(18), workbookPart),
                    PriceBetween25_49 = GetCellValue(cells.ElementAtOrDefault(19), workbookPart),
                    PriceBetween50_99 = GetCellValue(cells.ElementAtOrDefault(20), workbookPart),
                    PriceUpTo100 = GetCellValue(cells.ElementAtOrDefault(21), workbookPart),
                    VKValidSince = GetCellValue(cells.ElementAtOrDefault(22), workbookPart),
                    EKDelivery = GetCellValue(cells.ElementAtOrDefault(23), workbookPart),
                    EKDeliverySince = GetCellValue(cells.ElementAtOrDefault(24), workbookPart),
                    MCertificate = GetCellValue(cells.ElementAtOrDefault(25), workbookPart),
                };
                retrievalDataDto.Add(retrievalData);
            }

            return retrievalDataDto;
        }

        private static string GetCellValue(Cell? cell, WorkbookPart workbookPart)
        {
            if (cell is null || cell.CellValue is null || workbookPart.SharedStringTablePart is null)
            {
                return string.Empty;
            }

            var value = cell.CellValue.InnerText;

            if (cell.DataType is not null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                value = sharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }

            return value;
        }

        private static int GetColumnIndexFromCellReference(StringValue? cellReference)
        {
            if (cellReference is null)
            {
                return -1;
            }
            var reference = cellReference.Value;
            var columnIndex = 0;
            var factor = 1;

            for (var i = reference!.Length - 1; i >= 0; i--)
            {
                if (char.IsLetter(reference[i]))
                {
                    columnIndex += (reference[i] - 'A' + 1) * factor;
                    factor *= 26;
                }
            }
            return columnIndex - 1;
        }
    }
}
