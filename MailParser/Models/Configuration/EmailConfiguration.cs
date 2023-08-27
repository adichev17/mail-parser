namespace MailParser.Models.Configuration
{
    public class EmailConfiguration
    {
        public const string ConfigKey = "EmailConfiiguration";

        /// <summary>
        /// Your mail
        /// </summary>
        public string ImapServer { get; set; }

        /// <summary>
        /// Your Password 
        /// </summary>
        public string ImapUsername { get; set; }

        /// <summary>
        /// Adress of the SMTP host 
        /// </summary>
        public string ImapPassword { get; set; }

        /// <summary>
        /// Port of the host
        /// </summary>
        public int ImapPort { get; set; }

        /// <summary>
        /// Does the smtp host use SSL
        /// </summary>
        public bool ImapUseSSL { get; set; }
    }
}
