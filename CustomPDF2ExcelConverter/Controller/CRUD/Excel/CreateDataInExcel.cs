using CustomPDF2ExcelConverter.Controller.Enum;
using CustomPDF2ExcelConverter.Model;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CustomPDF2ExcelConverter.Controller
{
    public static class CreateDataInExcel
    {
        public static void InsertIntoExcelFile(SheetData sheetData, IReadOnlyCollection<RetrievalDataDto> retrievalData, IReadOnlyCollection<RetrievalDataDto> currentMovedRetrievalDataStateInExcel)
        {
            var startRowIndex = sheetData.Elements<Row>().Select(r => r.RowIndex!.Value).Max() + 1;

            foreach (var retrieval in retrievalData)
            {
                var ptsOrder = HasRequestBeenRecorded(retrieval, currentMovedRetrievalDataStateInExcel);
                string? status = null;
                
                if (!string.IsNullOrEmpty(ptsOrder))
                {
                    status = Status.ReEdit;
                }

                var newRow = new Row() { RowIndex = startRowIndex };

                newRow.Append(CreateCellForExcelOfType.TextCell("A", startRowIndex, retrieval.OrderNumber));
                newRow.Append(CreateCellForExcelOfType.TextCell("B", startRowIndex, ptsOrder ?? retrieval.PTSOrder));
                newRow.Append(CreateCellForExcelOfType.TextCell("C", startRowIndex, retrieval.Naming));
                newRow.Append(CreateCellForExcelOfType.TextCell("D", startRowIndex, retrieval.Plant));
                newRow.Append(CreateCellForExcelOfType.TextCell("E", startRowIndex, retrieval.UnloadingPoint));
                newRow.Append(CreateCellForExcelOfType.TextCell("F", startRowIndex, retrieval.ItemNumberCustomer));
                newRow.Append(CreateCellForExcelOfType.TextCell("G", startRowIndex, status ?? retrieval.Status));
                newRow.Append(CreateCellForExcelOfType.TextCell("H", startRowIndex, retrieval.LastDelivery));
                newRow.Append(CreateCellForExcelOfType.TextCell("I", startRowIndex, retrieval.WECaptureDate));

                newRow.Append(CreateCellForExcelOfType.TextCell("K", startRowIndex, retrieval.Appointment));
                newRow.Append(CreateCellForExcelOfType.TextCell("L", startRowIndex, retrieval.Quantity));
                newRow.Append(CreateCellForExcelOfType.TextCell("M", startRowIndex, retrieval.QuantityChange));

                sheetData.AppendChild(newRow);
                startRowIndex++;
            }
        }

        private static string? HasRequestBeenRecorded(RetrievalDataDto retrievalData, IReadOnlyCollection<RetrievalDataDto> currentMovedRetrievalDataStateInExcel)
        {
            if(currentMovedRetrievalDataStateInExcel is null)
            {
                return null;
            }

            return currentMovedRetrievalDataStateInExcel.FirstOrDefault(currentMovedData =>
                                currentMovedData.OrderNumber == retrievalData.OrderNumber &&
                                currentMovedData.ItemNumberCustomer == retrievalData.ItemNumberCustomer)?.PTSOrder;
        }
    }
}
