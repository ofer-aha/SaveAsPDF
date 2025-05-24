using System.Collections.Generic;
using System.Net.Mail;

namespace SaveAsPDF.Helpers
{
    public static class MailLogic
    {

        public static void SendEmail(string toAddress, string subject, string body)
        {
            SendEmail(new List<string> { toAddress }, new List<string>(), subject, body);
        }

        /// <summary>
        /// Sends an email to the specified recipients.
        /// </summary>
        /// <param name="toAddresses">The list of email addresses to send the email to.</param>
        /// <param name="bccAddresses">The list of email addresses to send the email as blind carbon copy (BCC) recipients.</param>
        /// <param name="subject">The subject of the email.</param>
        /// <param name="body">The body of the email.</param>
        public static void SendEmail(List<string> toAddresses, List<string> bccAddresses, string subject, string body)
        {
            MailAddress fromMailAddress = new MailAddress("senderEmail", "senderDisplayName");
            //TODO3:populate email details 

            MailMessage mail = new MailMessage();
            foreach (string toAddress in toAddresses)
            {
                mail.To.Add(toAddress);
            }
            foreach (string bccAddress in bccAddresses)
            {
                mail.To.Add(bccAddress);
            }
            mail.From = fromMailAddress;
            mail.Subject = subject;
            mail.Body = body;
            mail.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();

            client.Send(mail);
        }
    }

}

