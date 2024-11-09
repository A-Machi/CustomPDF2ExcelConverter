using CustomPDF2ExcelConverter.Controller.Enum;
using CustomPDF2ExcelConverter.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;
using Text = DocumentFormat.OpenXml.Spreadsheet.Text;

namespace CustomPDF2ExcelConverter.Controller
{
    public class UpdateDataInExcel
    {
        public static void UpdateDataFromWorksheet(SheetData sheetData, IReadOnlyList<RetrievalDataDto> oldData, IList<RetrievalDataDto> dataToReplaceOld)
        {
            foreach (var currentDataToUpdate in oldData)
            {
                var dataToUpdateWith = dataToReplaceOld.Where(updatedData => updatedData.OrderNumber == currentDataToUpdate.OrderNumber && updatedData.Appointment == currentDataToUpdate.Appointment).FirstOrDefault();
                var row = sheetData!.Elements<Row>().FirstOrDefault(r => r.RowIndex! == currentDataToUpdate.RowNumber);

                if (dataToUpdateWith is null || row is null)
                {
                    continue;
                }

                sheetData.RemoveChild(row);
                
                var changedQuantity = string.Empty;
                var diffQuantity = int.Parse(dataToUpdateWith.Quantity) - int.Parse(currentDataToUpdate.Quantity);
                if (!currentDataToUpdate.QuantityChange.Replace("+", "").Equals(diffQuantity.ToString()))
                {
                    changedQuantity = diffQuantity > 0 ? "+" + diffQuantity.ToString() : diffQuantity.ToString();
                }
                else
                {
                    changedQuantity = currentDataToUpdate.QuantityChange;
                }

                InsertDataIntoSheet(sheetData, dataToUpdateWith, currentDataToUpdate, changedQuantity);

                dataToReplaceOld.Remove(dataToUpdateWith);
            }
        }

        private static void InsertDataIntoSheet(SheetData sheetData, RetrievalDataDto dataToUpdateWith, RetrievalDataDto currentDataToUpdate, string changedQuantity)
        {
            var rowIndex = currentDataToUpdate.RowNumber;
            var row = new Row() { RowIndex = rowIndex };

            row.Append(CreateCellForExcelOfType.TextCell("A", rowIndex, currentDataToUpdate.OrderNumber));
            row.Append(CreateCellForExcelOfType.TextCell("B", rowIndex, currentDataToUpdate.PTSOrder));
            row.Append(CreateCellForExcelOfType.TextCell("C", rowIndex, dataToUpdateWith.Naming));
            row.Append(CreateCellForExcelOfType.TextCell("D", rowIndex, dataToUpdateWith.Plant));
            row.Append(CreateCellForExcelOfType.TextCell("E", rowIndex, dataToUpdateWith.UnloadingPoint));
            row.Append(CreateCellForExcelOfType.TextCell("F", rowIndex, dataToUpdateWith.ItemNumberCustomer));

            if (!currentDataToUpdate.Status.Equals(Status.New))
            {
                row.Append(CreateCellForExcelOfType.TextCell("G", rowIndex, Status.ReEdit));
            }
            else
            {
                row.Append(CreateCellForExcelOfType.TextCell("G", rowIndex, Status.New));
            }

            row.Append(CreateCellForExcelOfType.TextCell("H", rowIndex, dataToUpdateWith.LastDelivery));
            row.Append(CreateCellForExcelOfType.TextCell("I", rowIndex, dataToUpdateWith.WECaptureDate));
            row.Append(CreateCellForExcelOfType.TextCell("J", rowIndex, currentDataToUpdate.VNumber));
            row.Append(CreateCellForExcelOfType.TextCell("K", rowIndex, currentDataToUpdate.Appointment));
            row.Append(CreateCellForExcelOfType.TextCell("L", rowIndex, dataToUpdateWith.Quantity));
            
            var cellOfChangedQuantity = CreateCellForExcelOfType.TextCell("M", rowIndex, changedQuantity);
            SetChangedQuantityAndStyle(cellOfChangedQuantity, changedQuantity);
            row.Append(cellOfChangedQuantity);

            row.Append(CreateCellForExcelOfType.TextCell("N", rowIndex, currentDataToUpdate.Changes));
            row.Append(CreateCellForExcelOfType.TextCell("O", rowIndex, currentDataToUpdate.Link));
            row.Append(CreateCellForExcelOfType.TextCell("P", rowIndex, currentDataToUpdate.Remark));
            row.Append(CreateCellForExcelOfType.TextCell("Q", rowIndex, currentDataToUpdate.PriceBetween1_4));
            row.Append(CreateCellForExcelOfType.TextCell("R", rowIndex, currentDataToUpdate.PriceBetween5_9));
            row.Append(CreateCellForExcelOfType.TextCell("S", rowIndex, currentDataToUpdate.PriceBetween10_24));
            row.Append(CreateCellForExcelOfType.TextCell("T", rowIndex, currentDataToUpdate.PriceBetween25_49));
            row.Append(CreateCellForExcelOfType.TextCell("U", rowIndex, currentDataToUpdate.PriceBetween50_99));
            row.Append(CreateCellForExcelOfType.TextCell("V", rowIndex, currentDataToUpdate.PriceUpTo100));
            row.Append(CreateCellForExcelOfType.TextCell("W", rowIndex, currentDataToUpdate.VKValidSince));
            row.Append(CreateCellForExcelOfType.TextCell("X", rowIndex, currentDataToUpdate.EKDelivery));
            row.Append(CreateCellForExcelOfType.TextCell("Y", rowIndex, currentDataToUpdate.EKDeliverySince));
            row.Append(CreateCellForExcelOfType.TextCell("Z", rowIndex, currentDataToUpdate.MCertificate));

            InsertRow(sheetData, row);
        }

        private static void InsertRow(SheetData sheetData, Row newRow)
        {
            var rowIndex = newRow.RowIndex!.Value;
            Row? refRow = null;

            foreach (var row in sheetData.Elements<Row>())
            {
                if (row.RowIndex!.Value > rowIndex)
                {
                    refRow = row;
                    break;
                }
            }

            if (refRow == null)
            {
                sheetData.AppendChild(newRow);
            }
            else
            {
                sheetData.InsertBefore(newRow, refRow);
            }
        }

        private static void SetChangedQuantityAndStyle(Cell cell, string text)
        {
            cell.CellValue = new CellValue(text);
            cell.DataType = new EnumValue<CellValues>(CellValues.Number);

            var colorValue = TextColor.Black;

            if (text.StartsWith('+')) { colorValue = TextColor.Green; };
            if (text.StartsWith('-')) { colorValue = TextColor.Red; };

            var runProperties = new RunProperties(new Color { Rgb = new HexBinaryValue(colorValue) });

            var run = new Run();
            run.Append(runProperties);
            run.Append(new Text(text));

            var inlineString = new InlineString();
            inlineString.Append(run);

            cell.InlineString = inlineString;
            cell.DataType = new EnumValue<CellValues>(CellValues.InlineString);
        }
    }
}
