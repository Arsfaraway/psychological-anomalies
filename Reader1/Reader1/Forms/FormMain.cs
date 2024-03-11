// snfsgsd6pzxvwjUjpJi8 совсем новый
//PgNt19dATQQX6UKbXr47 новый
//i2qeYJ76kyp4pZUK3FJ6 старый
using System;
using System.Collections.Generic;

using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using OpenPop.Mime;
using OpenPop.Pop3;
using OpenPop.Common;

using Message = OpenPop.Mime.Message;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.IO;
using System.Threading;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using OfficeOpenXml;
using Reader1.Services;
using Reader1.Models.Configuration;
using Reader1.Storage;
using Reader1.Forms;

namespace Reader1
{
    public partial class FormMain : Form
    {
        public FormMain()
        {
            InitializeComponent();
        }
        public static void SendMessageSimple(string _from, string _to, string _fromcopy, string _subject, string _message, string _attachfilename)
        {
  
            SmtpClient _smtp = new SmtpClient("smtp.mail.ru", 25);//smtp.yandex.ru;587,25,2525
            _smtp.Credentials = new NetworkCredential("fa.nsk@mail.ru", "snfsgsd6pzxvwjUjpJi8");//username,password
            _smtp.EnableSsl = true;
            MailMessage _mail = new MailMessage();
            _mail.From = new MailAddress(_from);
            _mail.To.Add(_to);
    
            if (_fromcopy.Length !=0)
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
                Console.WriteLine("Message sent successfully!");
            }
            catch
            {
                Console.WriteLine("Error!Message not sent!");
            }
        }
        public static void SendMessage(Message _msg, string _message, string _attachfilename)
        {


            SmtpClient _smtp = new SmtpClient("smtp.mail.ru", 25);//smtp.yandex.ru;587,25,2525
            _smtp.Credentials = new NetworkCredential("fa.nsk@mail.ru", "snfsgsd6pzxvwjUjpJi8");//username,password
            _smtp.EnableSsl = true;

            MailMessage _mail = new MailMessage();
            _mail.From = new MailAddress("fa.nsk@mail.ru");
            // парсинг адресов в поле То
            var addrTo = _msg.Headers.To.ToArray();
            foreach (var ado in addrTo)
            {
                _mail.To.Add(new MailAddress(ado.Address));
            }

            var addrcopy = _msg.Headers.Cc.ToArray();
            foreach (var ado in addrcopy)
            {
               _mail.CC.Add(new MailAddress(ado.Address));
            }

            _mail.SubjectEncoding = Encoding.UTF8;
            _mail.BodyEncoding = Encoding.UTF8;
            _mail.Subject = _msg.Headers.Subject;
            _mail.Body = _message;
            if (_attachfilename.Length != 0)
            {
                Attachment _attach = new Attachment(_attachfilename, MediaTypeNames.Application.Octet);
                _mail.Attachments.Add(_attach);
            }
            try
            {
                _smtp.Send(_mail);
                Console.WriteLine("Message sent successfully!");
            }
            catch
            {
                Console.WriteLine("Error!Message not sent!");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {

            //+Просканировать почтовый ящик fa.nsk@mail.ru на предмет новых писем
            //+Сохранить прикрепленные письма в папку  C:\Users\vladi\Desktop\Дом\Agregator\In 
            // добавить  наименование файла название школы в формате : Наименование школы_. Перед номером класса 
            //+ Отправить ответное письмо адресатам следующего содержания "Коллеги, добрый день! Письмо принято в обработку"

            // Используем using чтобы соединение автоматически закрывалось
            int max_count = 50;
            string file_name;

            using (OpenPop.Pop3.Pop3Client client = new Pop3Client())
              {
                  // Подключение к серверу
                  client.Connect("pop.mail.ru", 995, true);

                // Аутентификация (проверка логина и пароля)
                client.Authenticate("fa.nsk@mail.ru", "snfsgsd6pzxvwjUjpJi8", AuthenticationMethod.UsernameAndPassword);
                if (client.Connected)
                  {
                      // Получение количества сообщений в почтовом ящике
                      int messageAllCount = client.GetMessageCount();
                      int messageCount = 0;
                      int answerCount  = 0;

                      // Выделяем память под список сообщений. Мы хотим получить все сообщения
                      List <Message> allMessages = new List<Message>(messageAllCount);

                    // Сообщения нумеруются от 1 до messageCount включительно
                    // Другим языком нумерация начинается с единицы
                    // Большинство серверов присваивают новым сообщениям наибольший номер (чем меньше номер тем старее сообщение)
                    // Т.к. цикл начинается с messageCount, то последние сообщения должны попасть в начало списка
                    for (int i = messageAllCount; (i > 0) && (max_count > 0); i--)
                      {
                        
                        Message message = client.GetMessage(i);
                        
                        string subject  = message.Headers.Subject; //заголовок
                        string date     = message.Headers.Date.ToString(); //Дата/Время
                        string from     = message.Headers.From.ToString(); //от кого
                        string body = "";
                       /*if (String.Compare(from,"fa.nsk@mail.ru") = 0 )
                        {*/
                            max_count--;
                            messageCount++;
                            allMessages.Add(message);

                            // ищем первую плейнтекст версию в сообщении
                            MessagePart mpPlain = message.FindFirstPlainTextVersion();

                            if (mpPlain != null)
                            {
                                Encoding enc = mpPlain.BodyEncoding;
                                body = enc.GetString(mpPlain.Body); //получаем текст сообщения
                            }

                            ListViewItem mes = new ListViewItem(new string[] { body });
                            listView1.Items.Add(mes);
                        //}
                      }
                      // создать контейнер answerMessages 
                     // Выделяем память под список сообщений. Мы хотим ответить на все полученные сообщения
                     List<Message> answerMessages = new List<Message>(messageCount);

                    //Ищем во всех письмах все вложения
                    for (int i = 0; i < messageCount; i++)
                      {
                          Message msg = allMessages[i];

                         // надо узнать как  установить фильтр на тему
                          var att = msg.FindAllAttachments();
                          foreach (var ado in att)
                          {
                            // проверяем на существование файла в паке IN
                            if (File.Exists(Path.Combine(System.IO.Path.Combine("C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In", ado.FileName))))
                            {
                                // файл существует .xlsx
                                file_name = ado.FileName;
                                int dl = file_name.Length - 5;
                                int pos = 0;
                                string postr;
                                do
                                {
                                    pos++;
                                    postr = String.Concat("(", string.Format("{0:0}", pos));
                                    postr = String.Concat(postr, ")");
                                    file_name = file_name.Substring(0, dl);
                                    file_name = String.Concat(file_name, postr);
                                    file_name = String.Concat(file_name, ".xlsx");
                                }
                                while (File.Exists(Path.Combine(System.IO.Path.Combine("C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In", file_name))));
                            }
                            else
                            {
                                // файл не существует
                                file_name = ado.FileName;

                            }
                            //сохраняем все найденные в письмах вложения
                            ado.Save(new System.IO.FileInfo(System.IO.Path.Combine("C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In", file_name)));
                          }
                        //если в письме были вложения, удаляем данное письмо
                        if (att.Count > 0)
                        {
                            // наполнить контейнер answerMessages 
                            answerMessages.Add(msg);
                            answerCount++;
                            // подумать о переносе в отдельную папку в почте IN
                         //   client.DeleteMessage(messageCount - i);
                        }                     
                      }

                    //организован цикл по контейнеру answerMessages отправка писем 
                  for (int i = 0; i < answerCount; i++)
                    {
                        SendMessage(answerMessages[i], "Коллеги, добрый день! Письмо принято в обработку", "");
                                               
                    }
                }
            }
             
        }
        string SubstringBetweenSymbols(string _str, char _preSymbol, char _postSymbol)
        {
            int? preSymbolIndex = null;
            int? postSymbolIndex = null;


            for (int i = 0; i < _str.Length; i++)
            {
                if (i == 0 && _preSymbol == char.MinValue)
                {
                    preSymbolIndex = -1;
                }
                if (_str[i] == _preSymbol && !(preSymbolIndex.HasValue && _preSymbol == _postSymbol))
                {
                    preSymbolIndex = i;
                }
                if (_str[i] == _postSymbol && preSymbolIndex.HasValue && preSymbolIndex != i)
                {
                    postSymbolIndex = i;
                }
                if (i == _str.Length - 1 && _postSymbol == char.MinValue)
                {
                    postSymbolIndex = _str.Length;
                }



                if (preSymbolIndex.HasValue && postSymbolIndex.HasValue)
                {
                    var result = _str.Substring(preSymbolIndex.Value + 1, postSymbolIndex.Value - preSymbolIndex.Value - 1);
                    return result;
                }
            }



            return "";
        }
        void AddDataSource(string _filesource, string _filereciver)
        {
            //    _filesource.SOURCEDATA -> _filereciver.SOURCEDATA 
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook WBSource; //рабочая книга откуда будем копировать данные  
            Excel.Workbook WBReciver; //рабочая книга куда будем копировать данные
            Excel.Worksheet xlShtSource; //лист Excel            
            Excel.Worksheet xlShtReciver; //лист Excel            

            WBSource = xlApp.Workbooks.Open(_filesource); //название файла Excel откуда будем копировать лист
            WBReciver = xlApp.Workbooks.Open(_filereciver); //название файла Excel куда будем копировать лист
            xlShtSource  = WBSource.Worksheets["SOURCEDATA"]; //название листа 
            xlShtReciver = WBReciver.Worksheets["SOURCEDATA"]; //название листа 

            //xlSht.Copy(After: xlWB2.Worksheets[xlWB2.Worksheets.Count]);  //сам процесс копирования листа из одного файла в другой 
            //xlSht.Copy(xlWB2.Worksheets["SOURCEDATA"]);
            //MessageBox.Show("Лист '" + xlSht.Name.ToString() + "' успешно скопирован", "Поиск", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //string email_psy = xlSht.Cells[2, i].Text;
            // пока xlSht.Cells[2, i].Text пустое Pos ++

            int Pos_source = 0;
            //Pos_source = WBSource.Worksheets["SOURCEDATA"].Cells[WBSource.Worksheets["SOURCEDATA"].Rows.Count, 52].End().Row;
            int Pos_reciver = 0;
            string tmpstr;
            //поиск позиции начала записи новых параметров
            do
            {
                tmpstr = WBReciver.Worksheets["SOURCEDATA"].Cells[++Pos_reciver, 1].Text;
            }
            while (tmpstr.Length != 0);
            //подсчет скко записей скопировать

            do
            {
                tmpstr = WBSource.Worksheets["SOURCEDATA"].Cells[++Pos_source, 1].Text;

            }
            while (tmpstr.Length != 0);
            Pos_source--;
            // копирование данных
            xlShtReciver.Range[xlShtReciver.Cells[Pos_reciver, 1], xlShtReciver.Cells[Pos_source + Pos_reciver, 52]] = xlShtSource.Range[xlShtSource.Cells[2, 1], xlShtSource.Cells[Pos_source + 2, 52]].Value;


            WBReciver.Close(true); //закрываем и сохраняем изменения в файле 2
            WBSource.Close();
            xlApp.Quit(); //закрываем Excel
        }



        private string GetMapingEmail(string pEmail)
        {
            string tmpEmail = pEmail;
            string result = pEmail;
            //= pEmail.ToUpper();

            switch (tmpEmail)
            {
                case "e.kulikova@gym.nsk.ru": result = "E.KULIKOVA@GYM2.NSK.RU"; break;
                case "e.kulikova@gym2.nsk.ru": result = "E.KULIKOVA@GYM2.NSK.RU"; break;
                case "vika11052008@gmail.com": result = "VIKA11052008@GMAIL.COM"; break;
                case "ann.yanchenko@yandex.ru": result = "ANNN.YANCHENKO@YANDEX.RU"; break;
                case "anna.yanchenko@yandex.ru": result = "ANNN.YANCHENKO@YANDEX.RU"; break;
                case "annn.yanchenko@yandex.ru": result = "ANNN.YANCHENKO@YANDEX.RU"; break;
                case "landochkina_@mail.ru": result = "LANDOCHKINA_A@MAIL.RU"; break;
                case "landochkina_a@mail.ru": result = "LANDOCHKINA_A@MAIL.RU"; break;
                case "vasikevitch@yandex.ru": result = "VASIKEVITCH@YANDEX.RU"; break;
                case "grigiriybales@gmail.com": result = "GRIGORIYBALES@GMAIL.COM"; break;
                case "grigoriybales@gmail.com": result = "GRIGORIYBALES@GMAIL.COM"; break;
                case "grigoriybales@mail.ru": result = "GRIGORIYBALES@GMAIL.COM"; break;
                case "verronika@ngs.ru": result = "VERRONIKA@NGS.RU"; break;
                case "elle_ona@mail.ru": result = "ELLE_ONA@MAIL.RU"; break;
                case "elle_one@mail.ru": result = result = "ELLE_ONA@MAIL.RU"; break;
                case "larisa.matvienko.76@mail.ru": result = "LARISA.MATVIENKO.76@MAIL.RU"; break;
                case "elenatokareva@gmail.com": result = "ELENATOKAREVA355@GMAIL.COM"; break;
                case "elenatokareva355@gmail.com": result = "ELENATOKAREVA355@GMAIL.COM"; break;
                case "elenatokoreva335@gmail.com": result = "ELENATOKAREVA355@GMAIL.COM"; break;
                case "veta0990@rambler.ru": result = "VETA0990@RAMBLER.RU"; break;
                case "smv002@mail.ru": result = "SMV002@MAIL.RU"; break;
                case "revkova_nataiya@mail.ru": result = "REVKOVA_NATALYA@MAIL.RU"; break;
                case "revkova_nataliya@mail.ru": result = "REVKOVA_NATALYA@MAIL.RU"; break;
                case "revkova_natalya@mail.ru": result = "REVKOVA_NATALYA@MAIL.RU"; break;
                case "revkova_nataya@mail.ru": result = "REVKOVA_NATALYA@MAIL.RU"; break;
                case "anohinan89@mail.ru": result = "N.YAT89@MAIL.RU"; break;
                case "n.yat@mail.ru": result = "N.YAT89@MAIL.RU"; break;
                case "n.yat89@mail.ru": result = "N.YAT89@MAIL.RU"; break;
                case "talantseva.iu@yandex.ru": result = "TALANTSEVA.IU@YANDEX.RU"; break;
                case "talantseva@yandex.ru": result = "TALANTSEVA.IU@YANDEX.RU"; break;
                case "velichkoim@mail.ru": result = "VELICHKOIM@MAIL.RU"; break;
                case "marina2811170@mail.ru": result = "MARINA28111970@MAIL.RU"; break;
                case "marina28111970@mail.ru": result = "MARINA28111970@MAIL.RU"; break;
                case "elizaveta-6897@mail.ru": result = "PSY.HELP177@GMAIL.COM"; break;
                case "L.ERMOLAEVA35@yandex.ru": result = "EA.ERMOLAEVA35@GMAIL.COM"; break;
                case "psi.help177@gmail.com": result = "PSY.HELP177@GMAIL.COM"; break;
                case "psy.help177@gmail.com": result = "PSY.HELP177@GMAIL.COM"; break;
                case "iug02061975@gmail.com": result = "LUG02061975@GMAIL.RU"; break;
                case "lug02061975@gmail.com": result = "LUG02061975@GMAIL.RU"; break;
                case "lug02061975@mail.com": result = "LUG02061975@GMAIL.RU"; break;
                case "lug02061975@mail.ru": result = "LUG02061975@GMAIL.RU"; break;
                case "lug02061987@gmail.com": result = "LUG02061975@GMAIL.RU"; break;
                case "lyg02061975@gmail.com": result = "LUG02061975@GMAIL.RU"; break;
                case "lygina@mail.ru": result = "LUG02061975@GMAIL.RU"; break;
                case "astasiakinsk@inbox.ru": result = "NASTASIAKINSK@INBOX.RU"; break;
                case "Nastasiakins@inbox.ru": result = "NASTASIAKINSK@INBOX.RU"; break;
                case "nastasiakinsk@inbox.ru": result = "NASTASIAKINSK@INBOX.RU"; break;
                case "nastasiakinsk@mail.ru": result = "NASTASIAKINSK@INBOX.RU"; break;
                case "nastasiakinst@inbox.ru": result = "NASTASIAKINSK@INBOX.RU"; break;
                case "olga.usova.81@mail.ru": result = "OLGA.USOVA.81@MAIL.RU"; break;
                case "varvarakulagin@yandex.ru": result = "VARVARAKULAGIN@YANDEX.RU"; break;
                case "ariwa78@mail.ru": result = "ARIWA78@MAIL.RU"; break;
                case "riwa78@mail.ru": result = "ARIWA78@MAIL.RU"; break;
                case "aav@s217.ru": result = "AAV@S217.RU"; break;
                case "kaa@217.ru": result = "KAA@S217.RU"; break;
                case "kaa@s21.ru": result = "KAA@S217.RU"; break;
                case "kaa@s217.ru": result = "KAA@S217.RU"; break;
                case "sps217@mail.ru": result = "KAA@S217.RU"; break;
                case "yla@s217.ru": result = "YLA@S217.RU"; break;
                case "n-get@mail.ru": result = "N-GET@MAIL.RU"; break;
                case "mga46@mail.ru": result = "VIKA.VICTORIJA@MAIL.RU"; break;
                case "viika.victorja@mail.ru": result = "VIKA.VICTORIJA@MAIL.RU"; break;
                case "vika.victorija@maail.ru": result = "VIKA.VICTORIJA@MAIL.RU"; break;
                case "vika.victorija@mail.ru": result = "VIKA.VICTORIJA@MAIL.RU"; break;
                case "vika.viktorija@mail.ru": result = "VIKA.VICTORIJA@MAIL.RU"; break;
                case "aytach.gadzhieva@mail.ru": result = "AYTACH.GADZHIEVA@MAIL.RU"; break;
                case "lady.kuhareva@mail.ru": result = "LADY.KUHAREVA@MAIL.RU"; break;
                case "Ledi.Kuhareva@mail.ru": result = "LADY.KUHAREVA@MAIL.RU"; break;
                case "ledy.kuhareva@mail.ru": result = "LADY.KUHAREVA@MAIL.RU"; break;
                case "Lydy.kuhareva@mail.ru": result = "LADY.KUHAREVA@MAIL.RU"; break;
                case "Lady.Kuhareva@mail.ru": result = "LADY.KUHAREVA@MAIL.RU"; break;
                case "prelovskay.2020@mail.ru": result = "PRELOVSKAYA.2020@MAIL.RU"; break;
                case "prelovskaya.2020@mail.ru": result = "PRELOVSKAYA.2020@MAIL.RU"; break;
                case "KI-Levchenko@yandex.ru": result = "KI-LEVCHENKO@YANDEX.RU"; break;
                case "KI-Levchnko@yandex.ru": result = "KI-LEVCHENKO@YANDEX.RU"; break;
                case "kl-levchenko@yandex.ru": result = "KI-LEVCHENKO@YANDEX.RU"; break;
                case "ovpsyholog@yandex.ru": result = "OVPSYHOLOG@YANDEX.RU"; break;
                case "OVpsyholog@yndex.ru": result = "OVPSYHOLOG@YANDEX.RU"; break;
                case "vasikevich@yandex.ru": result = "VASIKEVITCH@YANDEX.RU"; break;
                case "vasikevitch@yanex.ru": result = "VASIKEVITCH@YANDEX.RU"; break;
                case "vik11052008@gmail.com": result = "VIKA11052008@GMAIL.COM"; break;
                case "vida@ngs.ru":result = "VIDA@NGS.RU"; break;
                case "enken@sch130.ru": result = "ENKEN@SCH130.RU"; break;
                case "s-kaminskaya@inbox.ru":result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "s-Kaminskaya@inbox.ru":result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "yulya.tchebanowa@yandex.ru": result = "YULYA.TCHEBANOWA@YANDEX.RU"; break;
                case "nusya1993@list.ru": result = "NUSYA1993@LIST.RU"; break;
                case "nysya1993@list.ru": result = "NUSYA1993@LIST.RU"; break;
                case "zheniichkahromcova@mail.ru": result = "ZHENICHKAHROMCOVA@MAIL.RU"; break;
                case "zhenichkahromcova@mail.ru": result = "ZHENICHKAHROMCOVA@MAIL.RU"; break;
                case "zhenchkahromcova@mail.ru": result = "ZHENICHKAHROMCOVA@MAIL.RU"; break;
                case "henichkahromcova@mail.ru": result = "ZHENICHKAHROMCOVA@MAIL.RU"; break;
                case "ellenatokareva@gmail.com": result = "elenatokareva355@gmail.com"; break;
                case "Kaminckaya@inbox.ru": result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "s - kaminskaya@inboox.ru": result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "s-kaminskaya@indox.ru": result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "yulya.tchebanowa@mail.ru": result = "YULYA.TCHEBANOWA@YANDEX.RU"; break;
                case "yulia.tchebanowa@yandex.ru": result = "YULYA.TCHEBANOWA@YANDEX.RU"; break;
                case "ariva78@mail.ru": result = "ARIWA78@MAIL.RU"; break;
                case "vika110520008@gmail.com": result = "VIKA11052008@GMAIL.COM"; break;
                case "veronika@ngs.ru": result = "VERRONIKA@NGS.RU"; break;
                case "verrunira@ngs.ru": result = "VERRONIKA@NGS.RU"; break;

                case "0184077@mail.mail.ru": result = "0184077@MAIL.RU"; break;
                case "ira-glushkova@eandex.ru": result = "IRA-GLUSHKOVA@YANDEX.RU"; break;
                case "elsergeeva@yandex.ru": result = "ELGSERGEEVA@YANDEX.RU"; break;
                case "veta@rambler.ru": result = "VETA0990@RAMBLER.RU"; break;
                case "revkova_natalua@mail.ru": result = "REVKOVA_NATALYA@MAIL.RU"; break;
                case "s-kaminskaye@inbox.ru": result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "s.kaminskaya@inbox.ru": result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "S-Kamininskaya@inbox.ru": result = "S-KAMINSKAYA@INBOX.RU"; break;
                case "alin.menshikova@mail.ru": result = "ALIN.MENSHIKOVA@YANDEX.RU"; break;
                case "Lug02061975@gmail.com": result = "LUG02061975@GMAIL.COM"; break;
                case "PSY.HELP177@GMAIL.COM": result = "L.ERMOLAEVA35@YANDEX.RU"; break;
                case "kate_veselova@bk.ru": result  = "KATE_VESELOVA@BK.RU"; break;
                case "kat_veselova@bk.ru":   result = "KATE_VESELOVA@BK.RU"; break;
            }

            /*
             ledy.Kuhareva@mail.ru на Lady.Kuhareva@mail.ru
             
             */
            return result;
        }
        private void button2_Click(object sender, EventArgs e)
        {

            // +просканировать папку C:\Users\vladi\Desktop\Дом\Agregator\In
            //+Открыть каждый файл на предмет парсинга email педагога-психолога
            // +Осуществить маппинг правильного ЭП педагога-психолога private string GetMapingEmail(string pEmail)
            //+Сформировать наименование файла по формату: Текущее наименование файла + (email педагога-психолога) + .xlsx
            //+Проверить существование файла из предыдущего пункта в папке C:\Users\vladi\Desktop\Дом\Agregator\Data
            //+Если файл отсутствует, то необходимо скопировать файл C:\Users\vladi\Desktop\Дом\Agregator\Data\Pattern\pattern.xlsx в файл :Текущее наименование файла + (email педагога-психолога) 
            // +наполнить лист SOURCEDATA из текущего файла 
            //+Если файл существует, то необходимо его открыть и добавить в конец листа SOURCEDATA информацию из текущего файла
            //+Текущий файл из C:\Users\vladi\Desktop\Дом\Agregator\In переместить в папку C:\Users\vladi\Desktop\Дом\Agregator\Arhive
            string folder_source   = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In\\";
            string folder_receiver = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Data\\";
            string folder_arhive   = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Arhive\\";
            
            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_source, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {
                    string fileName = currentFile.Substring(folder_source.Length);
                    string result_filename;
                    string tmpname = fileName.Remove (fileName.LastIndexOf('_'));
                        
                    Excel.Workbook xlWB;
                    Excel.Worksheet xlSht;

                    Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
                                                                       //xlApp.Visible = true;
                    xlWB = xlApp.Workbooks.Open(currentFile); //открываем наш файл           
                    xlSht = xlWB.Worksheets["SOURCEDATA"]; //
                    string email_psy = GetMapingEmail(xlSht.Cells[2, 8].Text);
                    //закрытие Excel
                    xlWB.Close(false); //закрываем файл
                    xlApp.Quit();
                    // сформировать наименование результирующего файла педагога психолога 
                    result_filename = String.Concat(tmpname, "(");
                    result_filename = String.Concat(result_filename, email_psy);
                    result_filename = String.Concat(result_filename, ")");
                    result_filename = String.Concat(result_filename, ".xlsx");
                    // определение существования результирующего файла folder_receiver + result_filename
                    if (File.Exists(Path.Combine(folder_receiver, result_filename)))
                    {
                      // файл folder_receiver + result_filename существует, значит добавляем в файл
                      //Если файл существует, то необходимо его открыть и добавить в конец листа SOURCEDATA информацию из текущего файла
                      AddDataSource(Path.Combine(folder_source, fileName), Path.Combine(folder_receiver, result_filename));
                    }
                    else
                    { 
                       // файл folder_receiver + result_filename не существует, значит создаем файл
                      //Если файл отсутствует, то необходимо скопировать файл C:\Users\vladi\Desktop\Дом\Agregator\Data\pattern.xlsx в файл :Текущее наименование файла + (email педагога-психолога) 
                      //и наполнить лист SOURCEDATA из текущего файла
                      File.Copy(Path.Combine(folder_receiver,"pattern.xlsx"), Path.Combine(folder_receiver, result_filename));
                      AddDataSource(Path.Combine(folder_source, fileName), Path.Combine(folder_receiver, result_filename));
                    }
                    // перенос из каталога In в Arhive
                    Directory.Move(currentFile, Path.Combine(folder_arhive, fileName));

                    ListViewItem mes = new ListViewItem(new string[] { result_filename });
                    listView1.Items.Add(mes);

                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message});
                listView1.Items.Add(mes);
            }


        }

        private void button3_Click(object sender, EventArgs e)
        {
            // все файлы имеющие в названии (email педагога-психолога) перенести C:\Users\vladi\Desktop\Дом\Agregator\Data в C:\Users\vladi\Desktop\Дом\Agregator\Out
            // +прочитать все файлы в папке C:\Users\vladi\Desktop\Дом\Agregator\Out
            // +распарсить адрес отправки
            // сформировать письмо : "Добрый день, коллеги! Направляем вам обобщенные данные ФА за 4 четверть 2022/2023 учебного года по классам, находящимся под вашим кураторством и прошедшие системную проверку"
            // +прикрепить обрабатываемый файл
            // +отправить письмо
            // +переместить файл из C:\Users\vladi\Desktop\Дом\Agregator\Out в C:\Users\vladi\Desktop\Дом\Agregator\Arhive
            string folder_source   = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Out\\";
            string folder_receiver = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Arhive\\Total\\";
            string folder_data     = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Data\\";
            string send_filename;
            string text_email = "Добрый день, коллеги! Направляем вам обобщенные данные ФА за 4 четверть 2022/2023 учебного года по классам, находящимся под вашим кураторством и прошедшие системную проверку";
            string mail_magistr = "fa@magistr54.ru";
            string mail_test = "natasha22barn@mail.ru";
            //все файлы имеющие в названии (email педагога-психолога) перенести C:\Users\vladi\Desktop\Дом\Agregator\Data в C:\Users\vladi\Desktop\Дом\Agregator\Out
            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_data, "*).xlsx");

                foreach (string currentFile in xlsFiles)
                {
                    string fileName = currentFile.Substring(folder_data.Length);
                    Directory.Move(currentFile, Path.Combine(folder_source, fileName));
                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }


            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_source, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {
                    string fileName = currentFile.Substring(folder_source.Length);

                    string send_mail = SubstringBetweenSymbols(fileName, '(', ')');
                    Directory.Move(currentFile, Path.Combine(folder_receiver, fileName));

                    string tmpfile = Path.Combine(folder_receiver, fileName);
                    //SendMessageSimple("fa.nsk@mail.ru", send_filename, "", fileName, text_email, tmpfile);
                    SendMessageSimple("fa.nsk@mail.ru", send_mail, "", "Обобощенные данные по ФА (4_ЧЕТВЕРТЬ_2022-2023_УЧ.ГОД)", text_email, tmpfile);
                    // для магистра
                    SendMessageSimple("fa.nsk@mail.ru", mail_magistr, "","Обобощенные данные по ФА (4_ЧЕТВЕРТЬ_2022-2023_УЧ.ГОД)",  text_email, tmpfile);
                    // для проверки
                    // SendMessageSimple(string _from, string _to, string _fromcopy, string _subject, string _message, string _attachfilename)
                    //SendMessageSimple("fa.nsk@mail.ru", mail_test, "","Обобощенные данные по ФА (4_ЧЕТВЕРТЬ_2022-2023_УЧ.ГОД)",  text_email, tmpfile);
                    //
                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Формирование единого загрузочного файла по итогам четверти C:\Users\vladi\Desktop\Дом\Agregator\Arhive\Total
            // +просканировать папку C:\Users\vladi\Desktop\Дом\Agregator\Arhive\Total
            // +пернести лист SOURCEDATA из каждого открытого файла в результирующий файл: C:\Users\vladi\Desktop\Дом\Agregator\Out\result.xlsx
            string folder_data = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Data\\";
            string folder_arhive = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Arhive\\Total\\";
            string folder_source = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Out\\";
            string result_filename = "result.xlsx";
            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_arhive, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {

                    // определение существования результирующего файла folder_source + result_filename
                    if (File.Exists(Path.Combine(folder_source, result_filename)))
                    {

                        // файл folder_source + result_filename существует, значит добавляем в файл
                        //Если файл существует, то необходимо его открыть и добавить в конец листа SOURCEDATA информацию из текущего файла
                        AddDataSource(Path.Combine(folder_arhive, currentFile.Substring(folder_arhive.Length)), Path.Combine(folder_source, result_filename));
                    }
                    else
                    {
                        // файл folder_source + result_filename не существует, значит создаем файл
                        //Если файл отсутствует, то необходимо скопировать файл folder_data + pattrn.xlsx в  folder_source + result_filename 
                        //и наполнить лист SOURCEDATA из текущего файла
                        File.Copy(Path.Combine(folder_data, "pattern.xlsx"), Path.Combine(folder_source, result_filename));
                        AddDataSource(Path.Combine(folder_arhive, currentFile.Substring(folder_arhive.Length)), Path.Combine(folder_source, result_filename));
                    }

                    ListViewItem mes = new ListViewItem(new string[] { currentFile.Substring(folder_arhive.Length) });
                    listView1.Items.Add(mes);

                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {
            //исключите наличе файлов с (i), где i = [1...]
            //просканировать папку C:\Users\vladi\Desktop\Дом\Agregator\In
            // определить пустоту файла на данные: на листе SOURCEDTA присутствует только заголовок
            // если в файле отсутствуют данные, то копировать файл с наименованием !+обследуемый файл
            string folder_source = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In\\";
            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_source, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(currentFile);
                    Excel.Worksheet xlWorksheet = xlWorkbook.Worksheets["SOURCEDATA"];
                    Excel.Range xlRange = xlWorksheet.UsedRange;
                    xlApp.Visible = false;
                    int rowCount = xlRange.Rows.Count;
                    //закрытие Excel
                    xlWorkbook.Close(false); //закрываем файл
                    xlApp.Quit();


                    // определение пустоты файлов
                    if (rowCount <= 2)
                    {
                        ListViewItem mes = new ListViewItem(new string[] { currentFile.Substring(folder_source.Length) });
                        listView1.Items.Add(mes);
                    }

                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }

        }
        private string GetNameInstitution(Excel.Range pRange, string pPos)
        {
            string result = pPos;
            return result;
        }
        private void button6_Click(object sender, EventArgs e)
        {
            // + Открыть файл в папке C:\Users\vladi\Desktop\Дом\Agregator\In\Отсев\Статистика.xlsx
            // + Загрузить в массив реестр учереждений с листа "Учреждения" столбцы (IDINSTITUTION,	NAME)
            // Открыть файл в папке C:\Users\vladi\Desktop\Дом\Agregator\In\Отсев\reestr.xlsx
            // Удалить все записи на листе "reestr", кроме 1 записи
            // Просканировать папку C:\Users\vladi\Desktop\Дом\Agregator\In
            //   Открыть каждый файл на предмет парсинга Ф.И.О,email педагога-психолога,Ф.И.О,email классного руководителя, класс
            //   Сохранить информцию парсинга  в файле C:\Users\vladi\Desktop\Дом\Agregator\In\Отсев\reestr.xlsx
            // Закрыть файл в папке C:\Users\vladi\Desktop\Дом\Agregator\In\Отсев\Статистика.xlsx
            // Закрыть файл в папке C:\Users\vladi\Desktop\Дом\Agregator\In\Отсев\reestr.xlsx

            string folder_source = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In\\";
            string folder_receiver = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\In\\Отсев\\";

            string file_Stat   = folder_receiver + "Статистика1.xlsx";
            string file_result = folder_receiver + "reestr.xlsx";
            int    rowCount, Pos_Result = 2;

            // сбор информации по учреждениям
            Excel.Workbook  xls_Stat_Book;
            Excel.Worksheet xls_Stat_Sheet;
            Excel.Range     xls_Stat_Range;


            Excel.Application xls_Stat_App = new Excel.Application();
            xls_Stat_Book           = xls_Stat_App.Workbooks.Open(file_Stat); //открываем наш файл           
            xls_Stat_Sheet          = xls_Stat_Book.Worksheets["Учреждения"]; //
            xls_Stat_Range          = xls_Stat_Sheet.UsedRange;
            rowCount                = xls_Stat_Range.Rows.Count;
            xls_Stat_App.Visible    = false;



            
            // инициализация результирующего файла
            Excel.Workbook xls_Result_Book;
            Excel.Worksheet xls_Result_Sheet;

            Excel.Application xls_Result_App = new Excel.Application();
            xls_Result_Book  = xls_Result_App.Workbooks.Open(file_result); //
            xls_Result_Sheet = xls_Result_Book.Worksheets["REESTR"]; //
            xls_Result_Sheet.Select();

            xls_Result_Sheet.Cells[1, 1] = "Классный руководитель";
            xls_Result_Sheet.Cells[1, 2] = "Класс";
            xls_Result_Sheet.Cells[1, 3] = "Психолог-педагог";
            xls_Result_Sheet.Cells[1, 4] = "e-mail психолога-педагога";
            xls_Result_Sheet.Cells[1, 5] = "Учреждение";

            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_source, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {
                    string fileName = currentFile.Substring(folder_source.Length);
                    //string tmppos = fileName.Substring(fileName.IndexOf('_',1)+1, fileName.IndexOf('_', 2)-5);
                    //fileName.IndexOf('_', 2)- fileName.IndexOf('_', 1)-1
                    Excel.Workbook xlWB;
                    Excel.Worksheet xlSht;

                    Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
                                                                        //xlApp.Visible = true;
                    xlWB = xlApp.Workbooks.Open(currentFile); //открываем наш файл           
                    xlSht = xlWB.Worksheets["SOURCEDATA"]; //


                    string strall = String.Concat(xlSht.Cells[2, 13].Text, xlSht.Cells[2, 14].Text);

                    xls_Result_Sheet.Cells[Pos_Result, 1] = xlSht.Cells[2, 4].Text;
                    xls_Result_Sheet.Cells[Pos_Result, 2] = strall;
                    xls_Result_Sheet.Cells[Pos_Result, 3] = xlSht.Cells[2, 7].Text;
                    xls_Result_Sheet.Cells[Pos_Result, 4] = xlSht.Cells[2, 8].Text;
                    xls_Result_Sheet.Cells[Pos_Result, 5] = fileName;
                    xls_Result_Sheet.Cells[Pos_Result, 6] = fileName.IndexOf('_', 1);
                    xls_Result_Sheet.Cells[Pos_Result, 7] = fileName.IndexOf('_', 2);

                    //GetNameInstitution(xls_Stat_Range, tmppos);
                    Pos_Result++;


                    //закрытие Excel
                    xlWB.Close(false); //закрываем файл
                    xlApp.Quit();

                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }

            xls_Result_Book.Close(true); //закрываем файл
            xls_Result_App.Quit();

            xls_Stat_Range.Delete();
            xls_Stat_Book.Close(false);
            xls_Stat_App.Quit();

            /*string strall = "", strtmp;
            for (int iRet = 2; iRet <= rowCount; iRet++)
            {

                strtmp = xls_Stat_Range.Cells[iRet, 2].Text;
                strall = String.Concat(strall, strtmp);
                strall = String.Concat(strall, ";");
            }*/


            /*      for (int iRet = 2; iRet <= rowCount; iRet++)
          {
              //strtmp = xls_Stat_Sheet.Cells[iRet, 2].Text;
              strtmp = xls_Stat_Range.Cells[iRet, 1].Text;
              strall = String.Concat(strall, strtmp);
              strall = String.Concat(strall, ";");
          }
              // загрузить учреждения в массив
              //--------------------
              // Create a list of parts.
              //--------------------

              //закрытие Excel
              //xls_Stat_Range.Delete();*/


            /*
            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_source, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {
                    string fileName = currentFile.Substring(folder_source.Length);
                    string tmpname = fileName.Remove(fileName.LastIndexOf('_'));

                    Excel.Workbook xlWB;
                    Excel.Worksheet xlSht;

                    Excel.Application xlApp = new Excel.Application(); //создаём приложение Excel
                                                                       //xlApp.Visible = true;
                    xlWB = xlApp.Workbooks.Open(currentFile); //открываем наш файл           
                    xlSht = xlWB.Worksheets["SOURCEDATA"]; //
                    string email_psy = GetMapingEmail(xlSht.Cells[2, 8].Text);
                    //закрытие Excel
                    xlWB.Close(false); //закрываем файл
                    xlApp.Quit();
                    // сформировать наименование результирующего файла педагога психолога 
                    result_filename = String.Concat(tmpname, "(");
                    result_filename = String.Concat(result_filename, email_psy);
                    result_filename = String.Concat(result_filename, ")");
                    result_filename = String.Concat(result_filename, ".xlsx");
                    // определение существования результирующего файла folder_receiver + result_filename
                    if (File.Exists(Path.Combine(folder_receiver, result_filename)))
                    {
                        // файл folder_receiver + result_filename существует, значит добавляем в файл
                        //Если файл существует, то необходимо его открыть и добавить в конец листа SOURCEDATA информацию из текущего файла
                        AddDataSource(Path.Combine(folder_source, fileName), Path.Combine(folder_receiver, result_filename));
                    }
                    else
                    {
                        // файл folder_receiver + result_filename не существует, значит создаем файл
                        //Если файл отсутствует, то необходимо скопировать файл C:\Users\vladi\Desktop\Дом\Agregator\Data\pattern.xlsx в файл :Текущее наименование файла + (email педагога-психолога) 
                        //и наполнить лист SOURCEDATA из текущего файла
                        File.Copy(Path.Combine(folder_receiver, "pattern.xlsx"), Path.Combine(folder_receiver, result_filename));
                        AddDataSource(Path.Combine(folder_source, fileName), Path.Combine(folder_receiver, result_filename));
                    }
                    // перенос из каталога In в Arhive
                    Directory.Move(currentFile, Path.Combine(folder_arhive, fileName));

                    ListViewItem mes = new ListViewItem(new string[] { result_filename });
                    listView1.Items.Add(mes);

                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }
            */
        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Формирование единого загрузочного файла по итогам четверти C:\Users\vladi\Desktop\Дом\Agregator\Arhive\Total
            // +просканировать папку C:\Users\vladi\Desktop\Дом\Agregator\Arhive\Total
            // +пернести лист SOURCEDATA из каждого открытого файла в результирующий файл: C:\Users\vladi\Desktop\Дом\Agregator\Out\result.xlsx
            string folder_data = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Data\\";
            string folder_arhive = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Arhive\\Total\\";
            string folder_source = "C:\\Users\\vladi\\Desktop\\Дом\\Agregator\\Out\\";
            string result_filename = "result.xlsx";
            try
            {
                var xlsFiles = Directory.EnumerateFiles(folder_arhive, "*.xlsx");

                foreach (string currentFile in xlsFiles)
                {

                    Excel.Workbook xls_Stat_Book;
                    Excel.Worksheet xls_Stat_Sheet;
                    Excel.Range xls_Stat_Range;

                    Excel.Application xls_Stat_App = new Excel.Application();
                    xls_Stat_Book = xls_Stat_App.Workbooks.Open(currentFile); //открываем наш файл           
                    xls_Stat_Sheet = xls_Stat_Book.Worksheets["SOURCEDATA"]; //
                    xls_Stat_Range = xls_Stat_Sheet.UsedRange;
                    int rowCount = xls_Stat_Range.Rows.Count;
                    xls_Stat_App.Visible = false;
                    xls_Stat_Range.Delete();
                    xls_Stat_Book.Close(false);
                    xls_Stat_App.Quit();

                    if (rowCount >= 2001)
                    {
                        ListViewItem mes = new ListViewItem(new string[] { currentFile.Substring(folder_arhive.Length) });
                        listView1.Items.Add(mes);

                    }


                }
            }
            catch (Exception eSend)
            {
                ListViewItem mes = new ListViewItem(new string[] { eSend.Message });
                listView1.Items.Add(mes);
            }

        }

        private void buttonSettings_Click(object sender, EventArgs e)
        {

            var form2 = new FormSetting();
            //form2.Closed += (s, args) => this.Close();
            form2.ShowDialog();

            //if (ConfigurationService.IsConfigured())
            //{
            //    ListViewItem greetingItem = new ListViewItem(new string[] { "Ваши данные успешно загружены" });

            //    listView1.Items.Add(greetingItem);
            //}
            //else
            //{
            //    ListViewItem greetingItem = new ListViewItem(new string[] { "Не удалось загрузить данные" });

            //    listView1.Items.Add(greetingItem);
            //}
        }



        private void FormMain_Load(object sender, EventArgs e)
        {

        }
    }
}
 