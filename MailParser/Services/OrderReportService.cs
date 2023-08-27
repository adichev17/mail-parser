using MailParser.Services.interfaces;
using MailParser.Models;
using MailParser.Models.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

namespace MailParser.Services
{
    public class OrderReportService : IOrderReportService
    {
        private readonly IEmailReceiver _emailReceiver;
        private readonly IExcelGenerator _excelGenerator;
        private readonly ILogger<OrderReportService> _logger;
        private readonly ReportConfiguration _reportOptions;

        public OrderReportService(IEmailReceiver emailReceiver, ILogger<OrderReportService> logger,
            IExcelGenerator excelGenerator, IOptions<ReportConfiguration> reportOptions)
        {
            _emailReceiver = emailReceiver;
            _logger = logger;
            _excelGenerator = excelGenerator;
            _reportOptions = reportOptions.Value;
        }

        public async Task GenerateReport()
        {
            var messages = await _emailReceiver.ReceiveMessages(_reportOptions.FolderName);
            messages = messages.ToList();
            _logger.LogInformation($"Found {messages.Count()} messages...");

            var orders = new List<OrderModel>();
            int i = 1;
            foreach (var message in messages)
            {
                if (!message.Subject.Contains("meduza-shop.ru") ||
                    !message.Subject.Contains("Регистрация покупки")) continue;

                var messageBodyText = message.TextBody ?? "";
                var createdString = messageBodyText.GetBetween("покупки*", "\r").Trim();
                var amountString = messageBodyText.GetBetween("Сумма по чеку (без копеек)*", "\r").Trim();

                if (!DateTime.TryParse(createdString, out DateTime created))
                {
                    _logger.LogError("Error getting \"Created\"");
                    continue;
                }

                if (!decimal.TryParse(amountString, out decimal amount))
                {
                    _logger.LogError("Error getting \"Amount\"");
                    continue;
                }

                var order = new OrderModel()
                {
                    Surname = messageBodyText.GetBetween("Фамилия*", "\r").Trim(),
                    PhoneNumber = messageBodyText.GetBetween("Телефон*", "\r").Trim(),
                    Amount = amount,
                    Created = created
                };

                orders.Add(order);
                _logger.LogInformation($"{i++} out of {messages.Count()} successfully processed");
            }

            var reportExcel = _excelGenerator.Generate(orders);
            await File.WriteAllBytesAsync("Report.xlsx", reportExcel);
        }
    }
}