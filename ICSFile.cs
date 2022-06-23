using System.Net.Mail;
using System.Net;

namespace Utils.ICS
{
    public class ICSFile
    {
        private readonly string _From = "From Name"; // FROM NAME CONFIGURATIN
        private readonly string _EmailFrom = "test@gmail.com"; // EMAIL USERNAME
        private readonly string _EmailPw = "testpw"; // EMAIL PASSWORD
        private readonly string _SMTP = "smtp.gmail.com"; // EMAIL SMTP CONFIGURATION
        private readonly int _PORT = 587; // EMAIL PORT CONFIGURATION
        private readonly string _PATH = Path.GetFullPath("./"); // PATH WHERE FILE WILL BE CREATED

        MailMessage Email;

        public ICSFile()
        {
            Email = new MailMessage();
        }

        public void SendDateConfirmEmail(string TO, string MESSAGE, string SUBJECT, string path)
        {

            try
            {
                if (TO == null)
                    TO = "otheremail@email.com";
                Email = new MailMessage(_EmailFrom, TO, SUBJECT, MESSAGE);
     
                Email.Attachments.Add(new Attachment(path)); // ATTACH .ICS FILE

                Email.IsBodyHtml = true;
                Email.From = new MailAddress(_EmailFrom, _From);
                Email.CC.Add("copyemail@email.com"); // COPY TO

                SmtpClient SMTPMail = new SmtpClient(_SMTP);
                SMTPMail.EnableSsl = true; // ENABLE SSL EMAIL CONFIGURATION
                SMTPMail.UseDefaultCredentials = false;
                SMTPMail.Host = _SMTP;
                SMTPMail.Port = _PORT;
                SMTPMail.Credentials = new NetworkCredential(_EmailFrom, _EmailPw);

                SMTPMail.Send(Email);

                SMTPMail.Dispose();


            }
            catch(Exception ex)
            {
                
            }
        }

        public void CreateICS()
        {
                string EmailTo = "toemail@email.com"; // EMAIL TO BE SEND
                DateTime BeginDate = DateTime.Now; // WHEN EVENT STARTS
                DateTime EndDate = DateTime.Now.AddHours(3); // WHEN EVENT FINISH
                Directory.CreateDirectory(_PATH); // CREATE DIRECTORY WHERE FILES WILL BE SAVED
                string FileName =  EmailTo.Trim() + DateTime.Now.ToString("ddMMyyyyhhmm") + ".ics"; // CONFIGURATE FILENAME TO BE SAVED
                var path = Path.Combine(_PATH, FileName.Trim());
                using (StreamWriter sw = File.AppendText(path)) // START WRITING .ICS FILE
                {
                    String ics = @"BEGIN:VCALENDAR
BEGIN:VEVENT
UID:judestudio"+
                               "\nDTSTAMP:" + FormatICSDate(DateTime.Now) +
                               "\nDESCRIPTION:" + "EVENT DESCRIPTION" +
                               "\nSUMMARY:" + "EVENT SUMMARY" +
                               "\nDTSTART:" + FormatICSDate(BeginDate) +
                               "\nDTEND:" + FormatICSDate(EndDate) +
                               "\nLOCATION:" + "Place or Direction of event" +
                               "\nEND:VEVENT" +
                               "\nEND:VCALENDAR";
                    sw.WriteLine(ics);
                    sw.Close();
                }

                // HTML EMAIL CREATION
                var subject = $"ICS File Creation";
                var body = "<div>"
                           + $"<h5 style='font-weight:bold;'>Hi!</h5>"
                           + "<p>This is the ICS File method creation </p>"
                           + "</div>";

                SendDateConfirmEmail(EmailTo, body, subject, path); // SEND EMAIL WITH .ICS FILE
            
          
        }

        public string FormatICSDate(DateTime d)
        {
            d = d.AddHours(5);
            var year = d.Year;
            var month = d.Month.ToString("D2");
            var day = d.Day.ToString("D2");
            var hour = d.Hour.ToString("D2");
            var minute = d.Minute.ToString("D2");
            var seconds = d.Second.ToString("D2");

            return String.Concat(year,month,day,"T",hour,minute,seconds, "Z");
        }
        
        
    

    }
}