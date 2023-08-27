using MimeKit;

namespace MailParser.Services.interfaces
{
    public interface IEmailReceiver
    {
        public Task<IEnumerable<MimeMessage>> ReceiveMessages(string folderName);
    }
}
