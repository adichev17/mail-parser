using MailParser.Models.Configuration;
using MailParser.Services;
using MailParser.Services.interfaces;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

var loggerFactory = LoggerFactory.Create(
    builder => builder
        .AddConsole()
        .AddDebug()
        .SetMinimumLevel(LogLevel.Debug)
);
var logger = loggerFactory.CreateLogger<Program>();

IConfiguration configuration = new ConfigurationBuilder()
    .AddJsonFile("appSettings.json", optional: true, reloadOnChange: true)
    .AddEnvironmentVariables()
    .Build();

var serviceCollection = new ServiceCollection();
serviceCollection.AddLogging(configure =>
    {
        configure
            .AddConsole()
            .AddDebug()
            .SetMinimumLevel(LogLevel.Debug);
    })
    .AddTransient<EmailReceiver>()
    .AddTransient<OrderReportService>()
    .AddTransient<ExcelGenerator>();
serviceCollection.AddSingleton<IEmailReceiver, EmailReceiver>();
serviceCollection.AddSingleton<IExcelGenerator, ExcelGenerator>();
serviceCollection.AddSingleton<IOrderReportService, OrderReportService>();
serviceCollection.AddOptions<EmailConfiguration>().Bind(configuration.GetSection(EmailConfiguration.ConfigKey));
serviceCollection.AddOptions<ReportConfiguration>().Bind(configuration.GetSection(ReportConfiguration.ConfigKey));
var serviceProvider = serviceCollection.BuildServiceProvider();


logger.LogInformation("Start parsing...");

var orderReportService = serviceProvider.GetService<IOrderReportService>();
await orderReportService.GenerateReport();

logger.LogInformation("All done!");
Console.ReadLine();