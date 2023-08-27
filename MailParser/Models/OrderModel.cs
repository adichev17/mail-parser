namespace MailParser.Models
{
    public class OrderModel
    {
        public string Surname { get; set; }
        public string PhoneNumber { get; set; }
        public decimal Amount { get; set; }
        public DateTime Created { get; set; }
    }
}
