using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace Mail_Send
{
    class Mail
    {
        private string login;
        private string password;
        private string pathToFolder;
        private List<string> ListFilesName;
        //Логин, пароль и путь к папке с PDF
        public Mail(string _login, string _password, string _pathToFolder)
        {
            login = _login;
            password = _password;
            pathToFolder = _pathToFolder;
        }
        public void Start()
        {
            GetName();
        }
        //Получаем 
        private void GetName()
        {
            ListFilesName = Directory.GetFiles(pathToFolder, "*", SearchOption.AllDirectories).ToList();
        }
        private void SendMail()
        {
            MailAddress from = new MailAddress(login, "");
            MailAddress to = new MailAddress("aleksandr.nesterov@rena-solutions.com");
            MailMessage m = new MailMessage(from, to);
            m.Subject = "Тест";
            m.Body = "Письмо-тест 2 работы smtp-клиента";
            SmtpClient smtp = new SmtpClient("smtp.yandex.ru", 587);
            smtp.Credentials = new NetworkCredential(login, password);
            smtp.EnableSsl = true;
            smtp.Send(m);
            //await smtp.SendMailAsync(m);
            Console.WriteLine("Письмо отправлено");
        }

    }
}
