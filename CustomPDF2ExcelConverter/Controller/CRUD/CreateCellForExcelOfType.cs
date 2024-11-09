using DocumentFormat.OpenXml.Spreadsheet;

namespace CustomPDF2ExcelConverter.Controller
{
    public class CreateCellForExcelOfType
    {
        public static Cell TextCell(string header, uint index, string text)
        {
            var cell = new Cell
            {
                DataType = CellValues.String,
                CellReference = header + index
            };
            cell.CellValue = new CellValue(text);
            return cell;
        }
    }
}
