namespace MailParser.Models.Configuration
{
    public class EmailConfiguration
    {
        public const string ConfigKey = "EmailConfiiguration";

        /// <summary>
        /// Your mail
        /// </summary>
        public string SmtpServer { get; set; }

        /// <summary>
        /// Your Password 
        /// </summary>
        public string SmtpUsername { get; set; }

        /// <summary>
        /// Adress of the SMTP host 
        /// </summary>
        public string SmtpPassword { get; set; }

        /// <summary>
        /// Port of the host
        /// </summary>
        public int SmtpPort { get; set; }

        /// <summary>
        /// Does the smtp host use SSL
        /// </summary>
        public bool SmtpUseSSL { get; set; }
    }
}
