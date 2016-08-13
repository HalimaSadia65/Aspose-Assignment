using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Net;
using System.Net.Mail;

namespace Aspose_Assignment_App.Models
{
     public class SendEmail
     {
          public void sendEmail(string email)
          {
               MailMessage message = new MailMessage();
               string body = "";
               message.Body = body;
               message.From = new MailAddress("hrinfo020@gmail.com");
               message.To.Add(email);
               message.Subject = "Salary Increment Letter";
               Attachment attachment = new Attachment("D:\\Aspose_Assignment\\Output.docx");
               message.Attachments.Add(attachment);
               message.IsBodyHtml = true;

               SmtpClient smtp = new SmtpClient();
               smtp.Host = "smtp.gmail.com";
               System.Net.NetworkCredential ntwd = new NetworkCredential();
               ntwd.UserName = "hrinfo020@gmail.com";
               ntwd.Password = "020hrinfo";
               smtp.UseDefaultCredentials = true;
               smtp.Credentials = ntwd;
               smtp.Port = 587;
               smtp.EnableSsl = true;

               smtp.Send(message);

          }
     }
}