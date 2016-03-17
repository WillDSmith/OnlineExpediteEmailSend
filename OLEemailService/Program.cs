using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using Microsoft.Exchange.WebServices.Data;
using System.Configuration;

namespace OLEemailService
{
    class Program
    {
        static void Main(string[] args)
        {
            ssReports report = new ssReports();
            InternalEntities db = new InternalEntities();
            
            //string dtback = DateTime.Now.ToString("yyyy-MM-dd");
            //DateTime date = Convert.ToDateTime(dtback);

            DateTime dtt = DateTime.Now;
            dtt = dtt.AddHours(-12);
            
            string dt = DateTime.Now.ToString("MMddyyyyhhmmss");
            string path = @"ExportedFiles\";
            string filename = "OLER" + dt + ".xlsx";
            string log = @"Logs\";                                                             
            string logfilename = dt + ".txt";

            if (!Directory.Exists(log))
            {
                Directory.CreateDirectory(log);
            }

            FileStream filestream = new FileStream(log + logfilename, FileMode.Create);
            var streamwriter = new StreamWriter(filestream);
            streamwriter.AutoFlush = true;
            Console.SetOut(streamwriter);
            Console.SetError(streamwriter);

            try
            {
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                report.CreateExcelDoc(path + filename);

                streamwriter.WriteLine("Excel file has been created!");
            }

            catch (Exception ex)
            {
                streamwriter.WriteLine(ex.ToString());
            }
            
            try
            {
                String[] mEmailTo = ConfigurationManager.AppSettings["EmailTo"].ToString().Split(',');
                string mEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
                string mEmailSubject = ConfigurationManager.AppSettings["EmailSubject"];
                
                ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.UseDefaultCredentials = true;
                service.AutodiscoverUrl(mEmailFrom, RedirectionUrlValidationCallback);

                EmailMessage email = new EmailMessage(service);
                foreach(String s in mEmailTo)
                {
                    email.ToRecipients.Add(s);
                }
                
                email.Subject = mEmailSubject;
                email.Body = new MessageBody("This email contains attachments");
                email.Attachments.AddFileAttachment(path + filename);

                streamwriter.WriteLine("Sending email...");
                email.Send();
                streamwriter.WriteLine("Email Sent....!");

                var entries = db.OnlineExpedites.Where(d => d.CreationDate > dtt || d.DateSentTimeStamp == null);
                if (entries != null)
                {
                    foreach (var entry in entries)
                    {
                        entry.DateSentTimeStamp = DateTime.Now;
                    }
                    db.SaveChanges();
                }
               
                streamwriter.WriteLine("Database Updated.....");
            }
            catch (Exception ex)
            {
                streamwriter.WriteLine(ex.ToString());
            }
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
