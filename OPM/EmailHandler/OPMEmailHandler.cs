using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace OPM.EmailHandler
{
    class OPMEmailHandler
    {
        public static void fSendEmail(string htmlString)
        {
            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
            mail.From = new MailAddress("taduydoan.ansv@gmail.com");
            mail.To.Add("taduydoan@ansv.vn");
            mail.Subject = "Test Mail";
            mail.Body = "This is for testing SMTP mail from GMAIL";
            SmtpServer.Port = 587;
            SmtpServer.Credentials = new System.Net.NetworkCredential("taduydoan.ansv@gmail.com", "Nuocmat@172");
            SmtpServer.EnableSsl = true;
            SmtpServer.Send(mail);
        }
    }
}
