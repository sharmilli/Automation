using System;
using System.Collections.Generic;
using System.Net.Mail;
using System.Configuration;
using System.Threading.Tasks;

namespace TaskUtility.Utilities
{
    public static class EmailUtlity
    {
        private static string server = ConfigurationManager.AppSettings.Get("MailServer");
        private static int port = Convert.ToInt32(ConfigurationManager.AppSettings.Get("Port"));
        private static string serverMailId = ConfigurationManager.AppSettings.Get("Sender");
        private static string serverLogon = ConfigurationManager.AppSettings.Get("Password");
        public static bool SendEmail(string from,List<string> toList, List<string> ccList, string subject,string message)
        {
            try
            {
                MailMessage mail = new MailMessage();
                SmtpClient smtpServer = new SmtpClient(server, port);
                smtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;
                mail.From = new MailAddress(from);
                Parallel.ForEach(toList, e => mail.To.Add(e));
                Parallel.ForEach(ccList, e => mail.CC.Add(e));
                mail.Subject = subject;
                mail.Body = message;
                smtpServer.UseDefaultCredentials = false;
                smtpServer.Credentials = new System.Net.NetworkCredential(serverMailId, serverLogon);
                smtpServer.EnableSsl = true;
                smtpServer.Send(mail);
                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Email sending failed...."+ex.Message);
                return false;

            }
        }
    }
}
