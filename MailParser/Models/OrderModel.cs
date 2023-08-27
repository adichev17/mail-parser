namespace MailParser.Models
{
    public class OrderModel
    {
        public string Surname { get; set; }
        public string PhoneNumber { get; set; }
        public string Amount { get; set; }
        /// <summary>
        /// customer requirement, user form data
        /// </summary>
        public string Created { get; set; }
    }
}
