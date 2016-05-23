using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;
using System.Net;

namespace Lexiconer
{ 
    public class Mailer
    {
        public static void MailSend(string label, string s)
        {
            string _s = s;
            SmtpClient Smtp = new SmtpClient("smtp.", 25);
            //Smtp.Credentials = new NetworkCredential("login", "pass");
            Smtp.EnableSsl = true;
            MailMessage Message = new MailMessage();
            Message.From = new MailAddress("lexiconer@extremail.ru");
            Message.To.Add(new MailAddress("lexiconer@extremail.ru"));
            Message.Subject = "тема";
            Message.Body += label;
            Message.Body += _s;
            try
            {
                Smtp.Send(Message);
            }
            catch (SmtpException)
            {
                //Console.WriteLine("Ошибка!");
            }
        }

        public static void MailSend()
        {
            SmtpClient Smtp = new SmtpClient("smtp.", 25);
            //Smtp.Credentials = new NetworkCredential("login", "pass");
            Smtp.EnableSsl = true;

            MailMessage Message = new MailMessage();
            Message.From = new MailAddress("lexiconer@extremail.ru");
            Message.To.Add(new MailAddress("lexiconer@extremail.ru"));
            Message.Subject = "тема";
            Message.Body = "s";
            try
            {
                Smtp.Send(Message);
            }
            catch (SmtpException)
            {
                //Console.WriteLine("Ошибка!");
            }
        }
    }      
}
