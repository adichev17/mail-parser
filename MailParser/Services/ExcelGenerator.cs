using MailParser.Models;
using MailParser.Services.interfaces;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MailParser.Services
{
    public class ExcelGenerator : IExcelGenerator
    {
        private readonly ILogger<ExcelGenerator> _logger;
        public ExcelGenerator(ILogger<ExcelGenerator> logger)
        {
            _logger = logger;
        }
        public byte[] Generate(List<OrderModel> models)
        {
            _logger.LogInformation("Start report generation...");
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets
                .Add("Order Report");

            sheet.Cells[1, 1, 1, 5].LoadFromArrays(new object[][] { new[] { "№", "Фамилия", "Телефон", "Сумма по чеку ₽", "Дата покупки"} });
            var row = 2;
            var column = 1;
            int i = 1;
            foreach (var item in models)
            {
                sheet.Cells[row, column].Value = i++;
                sheet.Cells[row, column + 1].Value = item.Surname;
                sheet.Cells[row, column + 2].Value = item.PhoneNumber;
                sheet.Cells[row, column + 3].Value = item.Amount;
                sheet.Cells[row, column + 4].Value = item.Created;
                row++;

                sheet.Cells[1, 1, row, column + 4].AutoFitColumns();
                sheet.Column(1).Width = 16;
                sheet.Column(2).Width = 16;
                sheet.Column(3).Width = 16;
                sheet.Column(4).Width = 16;
                sheet.Column(5).Width = 16;

                sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[1, 1, 1 + models.Count(), 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                sheet.Cells[1, 1, 1, 5].Style.Font.Bold = true;

                sheet.Cells[1, 1, 1 + models.Count(), 5].Style.Border.BorderAround(ExcelBorderStyle.Double);
                sheet.Cells[1, 1, 1, 5].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            }

            _logger.LogInformation("Report generated");

            return package.GetAsByteArray();
        }
    }
}
