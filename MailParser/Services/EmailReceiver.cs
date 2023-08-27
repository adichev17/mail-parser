using MailKit;
using MailKit.Net.Imap;
using MailParser.Models;
using MailParser.Models.Configuration;
using MailParser.Services.interfaces;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using MimeKit;
using Org.BouncyCastle.Asn1.X509;
using System.Collections.Generic;

namespace MailParser.Services
{
    public class EmailReceiver : IEmailReceiver
    {
        private readonly EmailConfiguration _emailConfiguration;
        private readonly ILogger<EmailReceiver> _logger;

        public EmailReceiver(
            IOptions<EmailConfiguration> emailConfiguration,
            ILogger<EmailReceiver> logger)
        {
            _logger = logger;
            _emailConfiguration = emailConfiguration.Value;
        }

        public async Task<IEnumerable<MimeMessage>> ReceiveMessages(string folderName)
        {
            using var client = new ImapClient();
            try
            {
                await client.ConnectAsync(_emailConfiguration.ImapServer, (_emailConfiguration).ImapPort, true);
                client.AuthenticationMechanisms.Remove("XOAUTH2");
                await client.AuthenticateAsync(_emailConfiguration.ImapUsername, _emailConfiguration.ImapPassword);
                var folder = await client.GetFolderAsync(folderName);
                await folder.OpenAsync(FolderAccess.ReadOnly);

                return folder.ToList();
            }
            finally
            {
                await client.DisconnectAsync(true);
                client.Dispose();
            }
        }
    }
}
