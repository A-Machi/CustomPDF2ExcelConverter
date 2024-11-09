using CustomPDF2ExcelConverter.Model;
using FluentAssertions;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace CustomPDF2ExcelConverter.Controller.Test
{
    public class ComparisonCheckTests
    {

        private List<RetrievalDataDto> PDFData = new()
            {
                new ()
                {
                    OrderNumber = "123456",
                    Plant = "42",
                    UnloadingPoint = "040",
                    ItemNumberCustomer = "25.7-06741.01B",
                    WECaptureDate = "14/10/24",
                    Naming = "Right tailpipe",
                    LastDelivery = "-",
                    Appointment = "31/07/25",
                    Quantity = "80",
                },

                new ()
                {
                    OrderNumber = "7891011",
                    Plant = "43",
                    UnloadingPoint = "045",
                    ItemNumberCustomer = "30.9-16541.11P",
                    WECaptureDate = "15/10/25",
                    Naming = "Left tailpipe",
                    LastDelivery = "012345 / 25",
                    Appointment = "01/09/29",
                    Quantity = "800",
                },
            };

        private List<RetrievalDataDto> ExcelData = new()
            {
                new ()
                {
                    RowNumber = 0,
                    OrderNumber = "123456",
                    Plant = "42",
                    UnloadingPoint = "040",
                    ItemNumberCustomer = "25.7-06741.01B",
                    WECaptureDate = "14/10/24",
                    Naming = "Right tailpipe",
                    LastDelivery = "-",
                    Appointment = "31/07/25",
                    Quantity = "80",
                    Status = "Paused",
                    QuantityChange = "-5",
                    PTSOrder = "1234",
                    VNumber = "",
                    Changes = "",
                    Link = "",
                    Remark = "",
                    PriceBetween1_4 = "",
                    PriceBetween5_9 = "",
                    PriceBetween10_24 = "",
                    PriceBetween25_49 = "",
                    PriceBetween50_99 = "",
                    PriceUpTo100 = "",
                    VKValidSince = "",
                    EKDelivery = "",
                    EKDeliverySince = "",
                    MCertificate = "",
                },

                new ()
                {
                    RowNumber = 1,
                    OrderNumber = "7891011",
                    Plant = "43",
                    UnloadingPoint = "045",
                    ItemNumberCustomer = "30.9-16541.11P",
                    WECaptureDate = "15/10/25",
                    Naming = "Left tailpipe",
                    LastDelivery = "012345 / 25",
                    Appointment = "01/09/29",
                    Quantity = "800",
                    Status = "Running",
                    QuantityChange = "10",
                    PTSOrder = "5678",
                    VNumber = "10",
                    Changes = "TestChanged",
                    Link = "youTube.com",
                    Remark = "Cool App",
                    PriceBetween1_4 = "$5.99",
                    PriceBetween5_9 = "$5.99",
                    PriceBetween10_24 = "$5.99",
                    PriceBetween25_49 = "$5.99",
                    PriceBetween50_99 = "$5.99",
                    PriceUpTo100 = "$5.99",
                    VKValidSince = "2013",
                    EKDelivery = "$5.99",
                    EKDeliverySince = "2013",
                    MCertificate = "Test",
                },
            };

        [Fact]
        public void CompareData_ShouldReturnCorrectListOfDataToBeCreated()
        {
            var excelData = new List<RetrievalDataDto>();

            var (oldData, toReplaceOldData, toBeCreatedData, toBeMovedData) = ComparisonCheck.CompareData(PDFData, excelData);

            oldData.Should().BeEmpty();
            toBeMovedData.Should().BeEmpty();
            toReplaceOldData.Should().HaveCount(2);
            toBeCreatedData.Should().HaveCount(2);

            PDFData.OrderBy(x => x.Appointment);

            toBeCreatedData.Should().BeEquivalentTo(PDFData);
            toBeCreatedData.Should().BeInDescendingOrder(x => x.Appointment);
        }

        [Fact]
        public void CompareData_ShouldReturnCorrectListOfDataToBeUpdate()
        {
            PDFData.First().Plant = "1000";
            PDFData.First().Naming = "Middle";
            PDFData.First().UnloadingPoint = "040";
            PDFData.First().ItemNumberCustomer = "A123456-40";
            PDFData.First().LastDelivery = "31/10/24";
            PDFData.First().WECaptureDate = "31/11/24";

            var (oldData, toReplaceOldData, toBeCreatedData, toBeMovedData) = ComparisonCheck.CompareData(PDFData, ExcelData);

            toBeMovedData.Should().BeEmpty();
            toBeCreatedData.Should().BeEmpty();
            oldData.Should().HaveCount(1);
            toReplaceOldData.Should().HaveCount(1);

            PDFData.OrderBy(x => x.Appointment);

            oldData.Should().BeEquivalentTo([ExcelData.First()]);
            toReplaceOldData.Should().BeEquivalentTo([PDFData.First()]);
        }

        [Fact]
        public void CompareData_ShouldReturnCorrectListOfDataToBeMoved()
        {
            var modifiedPDF = PDFData.First();

            var (oldData, toReplaceOldData, toBeCreatedData, toBeMovedData) = ComparisonCheck.CompareData([modifiedPDF], ExcelData);

            toBeCreatedData.Should().BeEmpty();
            toReplaceOldData.Should().BeEmpty();
            oldData.Should().HaveCount(1);
            toBeMovedData.Should().HaveCount(1);

            var lastPartOfExcel = ExcelData.Last();
            oldData.Should().BeEquivalentTo([lastPartOfExcel]);
            toBeMovedData.Should().BeEquivalentTo([lastPartOfExcel]);
            toBeMovedData.Should().BeEquivalentTo(oldData);
        }
    }
}

