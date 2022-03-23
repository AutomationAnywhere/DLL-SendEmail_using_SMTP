using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace SendEmailSMTP
{
    public class SendEmailClass
    {
        public static string sendEmail(
            string fromAddress, 
            string toAddress,
            string ccAddress,
            string bccAddress,
            string subject,
            bool isHtmlBody,
            string body,
            string severHost, 
            int serverPort,
            bool secureConnection, 
            string username,
            string password,
            string attachments)
        {
            
            try
            {
              
                MailMessage message = new MailMessage();
                SmtpClient smtp = new SmtpClient();
                message.From = new MailAddress(fromAddress);
                if (toAddress.Length > 0)
                {
                    string[] toAdrs = toAddress.Split(',');
                    foreach (string toEmail in toAdrs)
                    {
                        message.To.Add(new MailAddress(toEmail));
                        LogToFile(" INFO : New Email added as TO Address : " + toEmail);
                    }
          
                }
                if (ccAddress.Length > 0)
                {
                    string[] CCId = ccAddress.Split(',');
                    foreach (string CCEmail in CCId)
                    {
                        message.CC.Add(new MailAddress(CCEmail)); //Adding Multiple CC email Id  
                        LogToFile(" INFO : New email address added to CC: " + CCEmail);
                    }

                }

                if (bccAddress.Length > 0)
                {
                    string[] BCCId = bccAddress.Split(',');
                    foreach (string BCCEmail in BCCId)
                    {
                        message.Bcc.Add(new MailAddress(BCCEmail)); //Adding Multiple BCC email Id  
                        LogToFile(" INFO : New email address added to bcc : " + BCCEmail);
                    }
                }


                message.Subject = subject;
                message.IsBodyHtml = isHtmlBody;
                message.Body = body;
                System.Net.Mail.Attachment attachment;
                if (attachments.Length > 0)
                {
                    string[] attachmentsArr = attachments.Split(',');
                    for (int i = 0; i < attachmentsArr.Length; i++)
                    {
                        attachment = new System.Net.Mail.Attachment(attachmentsArr[i]);
                        message.Attachments.Add(attachment);
                    }
                }
                else
                {
                    LogToFile(" INFO : No Attachments to be send");
                }
                smtp.Port =serverPort;
                smtp.Host = severHost;
                LogToFile(" INFO : Configured smtp host : " + severHost);
                LogToFile(" INFO : Configured smtp port : " + serverPort);
                smtp.EnableSsl = secureConnection;
                smtp.UseDefaultCredentials = false;
                LogToFile(" INFO : SSL enable " + secureConnection);
                if (username.Length == 0)
                {
                    LogToFile(" INFO : Username is empty");
                    return "Username is empty";
                }
                if (password.Length == 0)
                {
                    LogToFile(" INFO : Password is empty");
                    return "Password is empty";
                }
                smtp.Credentials = new NetworkCredential(username, password);
                smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                smtp.Send(message);
                LogToFile(" INFO : Email send successfully");
                return "Email send successfully";
            }
            catch (Exception ex)
            {
                LogToFile(" ERROR: "+ ex.Message);
                return ex.Message;
            }
        }



        private static void LogToFile(string line)
        {
            string logFile = "EmailSendSmtp.log";
           string logFileFolderPath = getLogFolderPath();

                DateTime dateNow = DateTime.Now;
            //2022 - Mar - 23 Wed 09:09:50.623
            string now = dateNow.ToString("yyyy-MMM-dd ddd HH:mm:ss");

            if (!File.Exists(logFileFolderPath+ logFile))
            {
                if (!Directory.Exists(logFileFolderPath))
                {
                    Directory.CreateDirectory(logFileFolderPath);
                }
                File.Create(logFileFolderPath + logFile).Dispose();

                using (TextWriter tw = new StreamWriter(logFileFolderPath + logFile, true))
                { 
                    tw.WriteLine(now + ": "+ line);
                }

            }
            else if (File.Exists(logFileFolderPath + logFile))
            {
                using (TextWriter tw = new StreamWriter(logFileFolderPath + logFile, true))
                {
                    tw.WriteLine(now + ": " + line);
                }
            }
        }

        private static string getLogFolderPath()
        {
            string AAInstallationPath = checkInstalled("Automation Anywhere Bot Agent");
            XmlDocument xml = new XmlDocument();
            string xmlFilePath = AAInstallationPath + @"config\nodemanager-logging.xml";
            xml.Load(xmlFilePath);
            //XmlNodeList nodelist = xml.SelectNodes("//Property[@name='logPath']");
            XmlNode node = xml.SelectSingleNode("//Property[@name='logPath']");
            string logpath = node.InnerText + "/";
            if (logpath.Contains("${env:PROGRAMDATA}"))
            {
                string pgrmDataEnv = Environment.GetEnvironmentVariable("PROGRAMDATA");
                logpath = logpath.Replace("${env:PROGRAMDATA}", pgrmDataEnv);
                logpath = logpath.Replace("\\", "/");
            }
            return logpath;
        }
        private static string checkInstalled(string findByName)
        {
            string displayName;
            string InstallPath;
            string registryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";

            //64 bits computer
            RegistryKey key64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);
            RegistryKey key = key64.OpenSubKey(registryKey);

            if (key != null)
            {
                foreach (RegistryKey subkey in key.GetSubKeyNames().Select(keyName => key.OpenSubKey(keyName)))
                {
                    displayName = subkey.GetValue("DisplayName") as string;
                    if (displayName != null && displayName.Contains(findByName))
                    {

                        InstallPath = subkey.GetValue("InstallLocation").ToString();

                        return InstallPath; //or displayName

                    }
                }
                key.Close();
            }

            return null;
        }
    }
}
