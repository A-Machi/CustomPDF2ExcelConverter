using CustomPDF2ExcelConverter.Model;
using System.Globalization;

namespace CustomPDF2ExcelConverter.Controller
{
    public class ComparisonCheck
    {
        public static (
            IReadOnlyList<RetrievalDataDto> oldData,
            IList<RetrievalDataDto> toReplaceOldData,
            IReadOnlyList<RetrievalDataDto> toBeCreatedData,
            IReadOnlyList<RetrievalDataDto> toBeMovedData)
            CompareData(IReadOnlyList<RetrievalDataDto> extractedTextFromPDF, IReadOnlyList<RetrievalDataDto> currentDataInExcel)
        {
            var oldData = currentDataInExcel.Where(excel => !extractedTextFromPDF.Any(pdf => AreEqual(pdf, excel)))
                .OrderBy(x => x.OrderNumber)
                .ThenBy(x => x.Appointment)
                .ToArray();

            var toReplaceOldData = extractedTextFromPDF.Where(pdf => !currentDataInExcel.Any(excel => AreEqual(pdf, excel)))
                .OrderBy(x => x.OrderNumber)
                .ThenBy(x => x.Appointment)
                .ToList();

            var toBeMovedData = currentDataInExcel.Where(excel => !extractedTextFromPDF.Any(pdf => excel.OrderNumber == pdf.OrderNumber && excel.Appointment == pdf.Appointment))
                .OrderBy(x => x.OrderNumber)
                .ThenBy(x => x.Appointment)
                .ToArray();

            var toBeCreatedData = extractedTextFromPDF.Where(pdf => !currentDataInExcel.Any(excel => pdf.OrderNumber == excel.OrderNumber && pdf.Appointment == excel.Appointment))
                .OrderBy(x => x.OrderNumber)
                .ThenBy(x => x.Appointment)
                .ToArray();

            return (oldData, toReplaceOldData, toBeCreatedData, toBeMovedData);
        }

        private static bool AreEqual(RetrievalDataDto pdf, RetrievalDataDto excel)
        {
            return pdf.Plant == excel.Plant &&
                   pdf.Naming == excel.Naming &&
                   pdf.UnloadingPoint == excel.UnloadingPoint &&
                   pdf.ItemNumberCustomer == excel.ItemNumberCustomer &&
                   pdf.WECaptureDate == excel.WECaptureDate &&
                   pdf.LastDelivery == excel.LastDelivery &&
                   pdf.Quantity == excel.Quantity;
        }
    }
}
