using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Reader1.Messages
{
    public class StartMessage
    {
        public static bool MailChecking(string _from, string _to, string _fromcopy, string _subject, string _message, string _attachfilename)
        {

            SmtpClient _smtp = new SmtpClient("smtp.gmail.com", 587);

            _smtp.Credentials = new NetworkCredential("arserm8@gmail.co", "wuzj pedu pphz zwmz");
            _smtp.EnableSsl = true;
            MailMessage _mail = new MailMessage();
            _mail.From = new MailAddress(_from);
            _mail.To.Add(_to);

            if (_fromcopy.Length != 0)
            {
                _mail.CC.Add(new MailAddress(_fromcopy));
            }
            _mail.SubjectEncoding = Encoding.UTF8;
            _mail.BodyEncoding = Encoding.UTF8;
            _mail.Subject = _subject;
            _mail.Body = _message;
            if (_attachfilename.Length != 0)
            {
                Attachment _attach = new Attachment(_attachfilename, MediaTypeNames.Application.Octet);
                _mail.Attachments.Add(_attach);
            }
            try
            {
                _smtp.Send(_mail);
                //var a = MessageBox.Show("Вам на почту был отправлен код верефикации, введите его", "Сообщение1", MessageBoxButtons.YesNoCancel);
                return true;
            }
            catch
            {
                //var a = MessageBox.Show("Не удалось отправить сообщение", "Ошибка", MessageBoxButtons.YesNoCancel);
                return false;
            }
        }

        static public bool IsValidEmail(string email)
        {
            var trimmedEmail = email.Trim();

            if (trimmedEmail.EndsWith("."))
            {
                return false; // suggested by @TK-421
            }
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == trimmedEmail;
            }
            catch
            {
                return false;
            }
        }

    }
}
