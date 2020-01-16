using System;
using System.Net;
using System.Net.Mail;
using System.IO;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Threading;
using System.Reflection;

namespace Subscriber
{
    class Program
    {
        static string FullPathToDir = "";

        static void Main(string[] args)
        {
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("     РОЗСИЛКА EMAIL");
            Console.WriteLine("");
            Console.WriteLine("");

            string FullPathToXMLFile = "";
            string FullPathToXSLTFile = "";
            string Email = "";

            if (args.Length == 3)
            {
                FullPathToXMLFile = args[0];
                FullPathToXSLTFile = args[1];
                Email = args[2];
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("Не заданi параметри");
                Console.WriteLine();
                Console.WriteLine("Потрiбно запустити програму з параметрами: ");
                Console.WriteLine(" --> 1. Повний шлях до ХМЛ файлу даних");
                Console.WriteLine(" --> 2. Повний шлях до XSLT шаблону");
                Console.WriteLine(" --> 3. Email або список email роздiлених комою");
                Console.WriteLine();
                Console.ReadLine();
                return;
            }

            if (!File.Exists(FullPathToXMLFile))
            {
                Console.WriteLine();
                Console.WriteLine("Не знайдений ХМЛ файл: " + FullPathToXMLFile);
                Console.ReadLine();
                return;
            }

            if (!File.Exists(FullPathToXSLTFile))
            {
                Console.WriteLine();
                Console.WriteLine("Не знайдений XSLT файл шаблону: " + FullPathToXSLTFile);
                Console.ReadLine();
                return;
            }

            //Папка де знаходиться ХМЛ, сюди будеть писатися логи
            FullPathToDir = Path.GetDirectoryName(FullPathToXMLFile);

            WriteWorkLog("-------------------------------------");

            List<string> EmailValidList = new List<string>();

            if (!String.IsNullOrEmpty(Email))
            {
                Console.WriteLine();
                Console.WriteLine(WriteWorkLog("Перевiрка email"));

                string[] EmailList = Email.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string emailItem in EmailList)
                {
                    string emailItemUpdate = emailItem.Trim();
                    if (!String.IsNullOrEmpty(emailItemUpdate))
                    {
                        bool IsValid = IsValidEmail(emailItemUpdate);

                        Console.WriteLine(WriteWorkLog(" -> " + emailItemUpdate + ": " + (IsValid ? "ok" : "no")));

                        int t = EmailValidList.IndexOf(emailItemUpdate);
                        if (IsValid && EmailValidList.IndexOf(emailItemUpdate) == -1)
                            EmailValidList.Add(emailItemUpdate);
                    }
                }

                if (EmailValidList.Count == 0)
                {
                    Console.WriteLine();
                    Console.WriteLine(WriteWorkLog("Список перевiрених email пустий"));
                    return;
                }
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine(WriteWorkLog("Не вказаний Email"));
                return;
            }

            string FullPathToHtmlResultFile = Path.Combine(Path.GetDirectoryName(FullPathToXMLFile), Path.GetFileNameWithoutExtension(FullPathToXMLFile) + ".html");

            //Трансформацiя ХМЛ даних в НТМЛ
            XslCompiledTransform xslt = new XslCompiledTransform();

            try
            {
                xslt.Load(FullPathToXSLTFile);
            }
            catch(Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(WriteWorkLog("Помилка при загрузцi шаблону: " + ex.Message));
                return;
            }

            try
            {
                xslt.Transform(FullPathToXMLFile,  FullPathToHtmlResultFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(WriteWorkLog("Помилка при трансформацiї даних: " + ex.Message));
                return;
            }

            xslt = null;

            string Subject = "";
            string Company = "";
            //string isClean = "";

            try
            {
                //Навiгатор
                XPathDocument xpDocChanges = new XPathDocument(FullPathToXMLFile);
                XPathNavigator xpDocChangesNavigator = xpDocChanges.CreateNavigator();

                //Вибiрка заголовку листа
                XPathNavigator NodeSubject = xpDocChangesNavigator.SelectSingleNode("/root/report/subject");
                if (NodeSubject != null)
                    Subject = NodeSubject.Value;

                //Вибiрка Назви компанiї (вiд кого лист буде приходити)
                XPathNavigator NodeCompany = xpDocChangesNavigator.SelectSingleNode("/root/report/company");
                if (NodeCompany != null)
                    Company = NodeCompany.Value;

                //Вибiрка чи потрібно видаляти файли
                //XPathNavigator NodeIsClean = xpDocChangesNavigator.SelectSingleNode("/root/report/isclean");
                //if (NodeIsClean != null)
                //    isClean = NodeIsClean.Value;

                xpDocChangesNavigator = null;
                xpDocChanges = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine();
                Console.WriteLine(WriteWorkLog("Помилка при вибiрцi даних з XML файлу: " + ex.Message));
                return;
            }

            //Інформація про SMTP сервери
            List<SmtpInfo> SmtpInfoList = new List<SmtpInfo>();

            string SmtpInfoXmlPath = GetPathSmtpInfoFile();

            if (SmtpInfoXmlPath != "")
            {
                //Зчитування з ХМЛ файлу
                ReadSmtpInfoFromXmlFile(ref SmtpInfoList, SmtpInfoXmlPath);
            }

            Console.WriteLine("");
            Console.WriteLine(WriteWorkLog("Вiдправка листiв: "));

            foreach (string email in EmailValidList)
            {
                foreach (SmtpInfo SmtpInfoItem in SmtpInfoList)
                {
                    Console.WriteLine(WriteWorkLog("SMTP server: " + SmtpInfoItem.SmtpServer));

                    if (SendMail(email, Subject, Company, FullPathToHtmlResultFile, SmtpInfoItem))
                        break;
                }
            }

            //Очистка
            //Clean(new string[] { FullPathToXMLFile, FullPathToHtmlResultFile });

            //Вибіркова очистка
            //if (isClean == "yes")
            //{
            //    Clean(new string[] { FullPathToXMLFile });
            //}

            Console.WriteLine(WriteWorkLog("Готово"));

            Thread.Sleep(2000);
        }

        /// <summary>
        /// Очистка
        /// </summary>
        /// <param name="files">масив файлів</param>
        static void Clean(string[] files)
        {
            foreach (string file in files)
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    WriteWorkLog(ex.Message);
                }
        }

        /// <summary>
        /// Перевiрка Email на коректнiсть
        /// </summary>
        /// <param name="strIn">Email</param>
        /// <returns></returns>
        static bool IsValidEmail(string strIn)
        {
            return Regex.IsMatch(strIn, @"^([\w-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([\w-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$");
        }

        /// <summary>
        /// Запис тексту в лог
        /// </summary>
        /// <param name="dir">Папка куди писати логи</param>
        /// <param name="msg">Текст</param>
        /// <returns>Повертає те саме повідомлення що пише в лог</returns>
        public static string WriteWorkLog(string msg)
        {
            try
            {
                StreamWriter sw = File.AppendText(Path.Combine(FullPathToDir, DateTime.Now.ToString("dd_MM_yyyy") + "_log.txt"));
                sw.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " - " + " - " + msg);
                sw.Flush();
                sw.Close();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }           

            return msg;
        }

        //
        //

        /// <summary>
        /// Вiдправка листа
        /// </summary>
        /// <param name="email">email</param>
        /// <param name="subject">subject</param>
        /// <param name="fullPathToHtmlBodyFile">Шлях до файлу даних</param>
        /// <param name="SmtpInfoItem">Конфігурація SMTP</param>
        public static bool SendMail(string email, string subject, string company, string fullPathToHtmlBodyFile, SmtpInfo SmtpInfoItem)
        {
            MailAddress fromMailAddress = new MailAddress(SmtpInfoItem.Email, company);
            MailAddress toMailAddress = new MailAddress(email);

            try
            {
                Console.Write(" -> " + email + ": ");

                using (MailMessage mail = new MailMessage(fromMailAddress, toMailAddress))
                using (SmtpClient client = new SmtpClient(SmtpInfoItem.SmtpServer, SmtpInfoItem.Port))
                {
                    mail.Subject = subject;
                    mail.IsBodyHtml = true;
                    mail.Body = File.ReadAllText(fullPathToHtmlBodyFile);
                    
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Timeout = 20000;
                    client.Credentials = new NetworkCredential(fromMailAddress.Address, SmtpInfoItem.Pass);
                    client.EnableSsl = true;

                    client.Send(mail);
                }

                Console.Write(WriteWorkLog(" send\n"));
                return true;
            }
            catch (Exception e)
            {
                Console.Write(WriteWorkLog(" error: " + e.Message + "\n"));
                return false;
            }
        }

        /// <summary>
        /// Функція повертає шлях до файлу конфігурації SMTP
        /// </summary>
        /// <returns></returns>
        static string GetPathSmtpInfoFile()
        {
            string DirectoryAssembly = Path.GetDirectoryName(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            string DirectoryParent = Directory.GetParent(DirectoryAssembly).FullName;

            string SmtpInfoXmlPath = Path.Combine(DirectoryParent, "SmtpInfo.xml");

            if (File.Exists(SmtpInfoXmlPath))
            {
                return SmtpInfoXmlPath;
            }
            else
            {
                Console.WriteLine("Не знайдений файл конфігурації SmtpInfo.xml. Перевірте шлях: " + SmtpInfoXmlPath);
                Console.ReadLine();

                return "";
            }
        }

        /// <summary>
        /// Загрузка інформації з XML файлу
        /// </summary>
        /// <param name="path"></param>
        public static void ReadSmtpInfoFromXmlFile(ref List<SmtpInfo> SmtpInfoList, string path)
        {
            XPathDocument xpathDoc = new XPathDocument(path);
            XPathNavigator xpathDocNavigator = xpathDoc.CreateNavigator();

            XPathNodeIterator nodesSmtpInfo = xpathDocNavigator.Select("/root/smtpinfo");
            while (nodesSmtpInfo.MoveNext())
            {
                XPathNavigator nodeCurrent = nodesSmtpInfo.Current;

                SmtpInfo SmtpInfoItem = new SmtpInfo();
                SmtpInfoItem.SmtpServer = nodeCurrent.SelectSingleNode("server").Value;
                SmtpInfoItem.Email = nodeCurrent.SelectSingleNode("email").Value;
                SmtpInfoItem.Pass = nodeCurrent.SelectSingleNode("pass").Value;
                SmtpInfoItem.Port = int.Parse(nodeCurrent.SelectSingleNode("port").Value);

                SmtpInfoList.Add(SmtpInfoItem);
            }
        }
    }
}
