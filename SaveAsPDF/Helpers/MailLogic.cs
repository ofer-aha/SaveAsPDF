﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SaveAsPDF.Helpers
{
    public static class MailLogic
    {
        public static class EmailLogic
        {
            

            public static void SendEmail(string toAddress, string subject, string body)
            {
                SendEmail(new List<string> { toAddress }, new List<string>(), subject, body);
            }

            public static void SendEmail(List<string> toAddresses, List<string> bccAddresses, string subject, string body)
            {
                MailAddress fromMailAddress = new MailAddress("senderEmail", "senderDisplayName");
                //TODO:populate email details 
                
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
}
