using CustomPDF2ExcelConverter.Model;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf;
using System.Text.RegularExpressions;
using CustomPDF2ExcelConverter.Controller.Enum;
namespace CustomPDF2ExcelConverter.Controller
{
    public static class ReadDataFromPDF
    {
        public static IReadOnlyList<RetrievalDataDto> ExtractTextFromPDF(string pdfFilePath)
        {
            using (PdfReader reader = new PdfReader(pdfFilePath))
            using (PdfDocument pdfDoc = new PdfDocument(reader))
            {
                var allConvertedData = new List<RetrievalDataDto>();
                for (int pageNumber = 1; pageNumber <= pdfDoc.GetNumberOfPages(); pageNumber++)
                {
                    var textPerPage = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(pageNumber));
                    var convertedData = ConvertFromPDFToDesiredFormat(textPerPage);
                    allConvertedData = [.. allConvertedData, .. convertedData];
                }
                return allConvertedData;
            }
        }

        private static IList<RetrievalDataDto> ConvertFromPDFToDesiredFormat(string input)
        {
            var lines = input.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var data = new Dictionary<string, string>();
            var listOfAppointmentsAndQuanitiesAndStatus = new List<(string, string, string)>();
            var retrievalData = new List<RetrievalDataDto>();

            var didReachedAppointmentField = false;
            var previousLine = string.Empty;
            var pattern = $@":\s*(\S+)";

            foreach (string line in lines)
            {
                if (line.Contains("Bestellnummer:"))
                {
                    data["Bestellnummer"] = ExtractValueWithProvidedPattern(line, "Bestellnummer" + pattern);
                }
                else if (line.Contains("Werk:"))
                {
                    data["Werk"] = ExtractValueWithProvidedPattern(line, "Werk" + pattern);
                }
                else if (line.Contains("Abladestelle:"))
                {
                    data["Abladestelle"] = ExtractValueWithProvidedPattern(line, "Abladestelle" + pattern);
                }
                else if (line.Contains("Sachnummer Kunde:"))
                {
                    data["Sachnummer_Kunde"] = ExtractValueWithProvidedPattern(line, "Sachnummer Kunde" + pattern);
                }
                else if (line.Contains("Letzte Lieferung"))
                {
                    data["Letzte_Lieferung"] = ExtractValueWithProvidedPattern(line, $@"Letzte Lieferung:\s*([\d\s/]+)");
                    if (previousLine != null &&
                        Regex.IsMatch(previousLine, @"\d{2}\.\d{2}\.\d+"))
                    {
                        data["WE_Erfassungsdatum"] = previousLine.Trim();
                    }
                }
                else if (line.Contains("bez1"))
                {
                    data["bez1"] = ExtractValueWithProvidedPattern(line, "bez1=([^;=]+)(;|=)");
                }
                else if (line.Contains("Termin:"))
                {
                    didReachedAppointmentField = true;
                }
                else if (didReachedAppointmentField)
                {
                    var appointment = string.Empty;
                    var quantity = "0";
                    var status = Status.New;

                    var patternForQuantity = @"\s(\d+)\s(\d+)";
                    var findQuantity = Regex.Match(line, patternForQuantity);
                    if (findQuantity.Success)
                    {
                        quantity = findQuantity.Groups[1].Value.Trim();
                    }

                    var patternWithRangeDate = @"\d{2}\.\d{2}\.\d+\s-\s(\d{2}\.\d{2}\.\d+)";
                    var matchWithRangeDate = Regex.Match(line, patternWithRangeDate);
                    if (matchWithRangeDate.Success && !quantity.Equals("0"))
                    {
                        appointment = matchWithRangeDate.Groups[1].Value.Trim();
                        listOfAppointmentsAndQuanitiesAndStatus.Add((appointment, quantity, status));
                    }

                    var patternSingleDate = @"[A-Z]\s(\d{2}\.\d{2}\.\d+)";
                    var matchSingleDate = Regex.Match(line, patternSingleDate);
                    if (matchSingleDate.Success && !matchWithRangeDate.Success && !quantity.Equals("0"))
                    {
                        appointment = matchSingleDate.Groups[1].Value.Trim();
                        listOfAppointmentsAndQuanitiesAndStatus.Add((appointment, quantity, status));
                    }

                    var patternResidue = "33.33.33";
                    var matchResidue = Regex.Match(line, patternResidue);
                    if (matchResidue.Success && !quantity.Equals("0"))
                    {
                        appointment = DateTime.UtcNow.ToString("dd.MM.yy");
                        listOfAppointmentsAndQuanitiesAndStatus.Add((appointment, quantity, "RÜCKSTAND"));
                    }

                }

                previousLine = line;
            }

            foreach (var appointmentAndQuantityAndStatus in listOfAppointmentsAndQuanitiesAndStatus)
            {
                
                retrievalData.Add(
                    MapToDto(
                        data, 
                        appointmentAndQuantityAndStatus.Item1, 
                        appointmentAndQuantityAndStatus.Item2, 
                        appointmentAndQuantityAndStatus.Item3)
                    );
            }

            return retrievalData;
        }

        private static string ExtractValueWithProvidedPattern(string input, string pattern)
        {
            var match = Regex.Match(input, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return string.Empty;
        }

        private static RetrievalDataDto MapToDto(Dictionary<string, string> extractedData, string appointment, string quantity, string status) => new()
        {
            OrderNumber = extractedData["Bestellnummer"],
            Plant =  extractedData["Werk"],
            UnloadingPoint = $"Turmfeld \"{extractedData["Abladestelle"]}\"",
            ItemNumberCustomer = extractedData["Sachnummer_Kunde"],
            WECaptureDate = extractedData["WE_Erfassungsdatum"].Replace(".", newValue: "/"),
            Naming = extractedData["bez1"],
            LastDelivery = extractedData["Letzte_Lieferung"] == "/ " ? "-" : extractedData["Letzte_Lieferung"].Replace(".", newValue: "/"),
            Appointment = appointment.Replace(".", newValue: "/"),
            Quantity = quantity,
            Status = status,
        };
    }
}
