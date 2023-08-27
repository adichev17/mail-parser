using MailParser.Services.interfaces;
using MailParser.Models;
using MailParser.Models.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using HtmlAgilityPack;

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

                if (string.IsNullOrWhiteSpace(message.HtmlBody))
                {
                    _logger.LogError($"HtmlBody - {message.HtmlBody}");
                    continue;
                }

                var doc = new HtmlDocument();
                doc.LoadHtml(message.HtmlBody);
                var tableNode = doc.DocumentNode.SelectNodes("//table")[1];
                var trNodes = tableNode.SelectNodes("//tr").Skip(1);
                var values = trNodes
                    .Where(trNode =>
                        trNode.ChildNodes.Count > 3 && !string.IsNullOrWhiteSpace(trNode.ChildNodes[3].InnerText.Trim()))
                    .Select(trNode => trNode.ChildNodes[3].InnerText.Trim()).ToList();

                if (values.Count != 4)
                {
                    _logger.LogError($"Error parse HTML Nodes. Message Id - {message.MessageId}");
                    continue;
                };

                var order = new OrderModel
                {
                    Surname = values[0],
                    PhoneNumber = values[1],
                    Amount = values[2],
                    Created = values[3]
                };

                orders.Add(order);
                _logger.LogInformation($"{i++} out of {messages.Count()} successfully processed");
            }

            if (i == messages.Count())
            {
                _logger.LogInformation($"All messages successfully processed");
            }
            else
            {
                _logger.LogInformation($"Successfully processed messages: {i}");
                _logger.LogInformation($"Error processed messages: {messages.Count() - i}");
            }

            var reportExcel = _excelGenerator.Generate(orders);
            await File.WriteAllBytesAsync("Report.xlsx", reportExcel);
        }
    }
}