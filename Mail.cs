using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToPdf
{
    class Mail
    {
        private string login;
        private string password;
        private string pathToFolder;
        //Логин, пароль и путь к папке с PDF
        public Mail(string _login, string _password, string _pathToFolder)
        {
            login = _login;
            password = _password;
            pathToFolder = _pathToFolder;
        }
        //Получаем 
        public void GetName()
        {

        }
        public void SendMail()
        {
            MailAddress from = new MailAddress("_login", "_password");
            MailAddress to = new MailAddress("somemail@yandex.ru");
            MailMessage m = new MailMessage(from, to);
            m.Subject = "Тест";
            m.Body = "Письмо-тест 2 работы smtp-клиента";
            SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587);
            smtp.Credentials = new NetworkCredential("somemail@gmail.com", "mypassword");
            smtp.EnableSsl = true;
            //await smtp.SendMailAsync(m);
            Console.WriteLine("Письмо отправлено");
        }

    }
}
