using MailParser.Models;

namespace MailParser.Services.interfaces
{
    public interface IExcelGenerator
    {
        byte[] Generate(List<OrderModel> models);
    }
}
