using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mime;

namespace eMailSender
{
    public class MailSender
    {
        public MailSender()
        { }
        public MailSender(String mail, String server, String password, int port)
        {
            Mail = mail;
            Server = server;
            MailPassword = password;
            Port = port;
        }

        /// <summary>
        /// Mail account for sending
        /// </summary>
        public String Mail { set; get; }

        /// <summary>
        /// Server smtp for the mail account
        /// </summary>
        public String Server { set; get; }

        /// <summary>
        /// Password for mail account
        /// </summary>
        public String MailPassword { set; get; }

        /// <summary>
        /// Server port for smtp
        /// </summary>
        public int Port { set; get; }

        /// <summary>
        /// Send a mail  
        /// </summary>
        /// <param name="Senders">Senders for the mail</param>
        /// <param name="SendersCC">Senders cc for the mail</param>
        /// <param name="SendersBCC">Senders bcc for the mail</param>
        /// <param name="Subject">Subject for the mail</param>
        /// <param name="Body">Body of mail</param>
        /// <param name="ImagesAttached">List of images attached (only for HTML format)</param>
        /// <param name="FilesAttached">list of files attached to the mail</param>
        /// <param name="SSL">Is SSL required? </param>
        /// <param name="AuthenticationRequiered">The Authentication is required?</param>
        /// <param name="isHTML">The body of mail is HTML?</param>
        /// <param name="Result">Result of sending</param>
        /// <param name="mailPriority">Priority of mail</param>
        /// <returns>return true if the mail is submitted successfully</returns>
        public Boolean SendMail(List<String> Senders, List<String> SendersCC, List<String> SendersBCC, String Subject, String Body, List<ImageAttached> ImagesAttached, List<String> FilesAttached, Boolean SSL, Boolean AuthenticationRequiered,Boolean isHTML ,out String Result, System.Net.Mail.MailPriority mailPriority = System.Net.Mail.MailPriority.Normal)
        {
            Boolean Res = false;
            String Sender = String.Empty;
            Result = String.Empty;
            System.Net.Mail.MailMessage mailmsg = PreparateMailMessage(Senders, SendersCC, SendersBCC, Subject, Body, ImagesAttached, isHTML,out Sender);

            if (FilesAttached != null)
            {
                FilesAttached.ForEach(fileattached =>
                {
                    if (System.IO.File.Exists(fileattached))
                        mailmsg.Attachments.Add(new System.Net.Mail.Attachment(fileattached));
                });
            }
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient();
            smtp.Host = Server;
            if (AuthenticationRequiered)
                smtp.Credentials = new System.Net.NetworkCredential(Mail, MailPassword);
            smtp.Port = Port;
            smtp.EnableSsl = SSL;
            try
            {
                smtp.Send(mailmsg);
                Result = "Submitted successfully";
                Res = true;
            }
            catch (Exception ex)
            {
                String sFilesAttached = String.Empty;
                if (FilesAttached != null)
                    FilesAttached.ForEach(fileattached => sFilesAttached += fileattached + ", ");
                
                String Source = String.Empty;
                if (ex.Source != null)
                    Source = ex.Source;

                Result = "Error:" + ex.Message + " object:" + Source + " Remitente:" + Sender + " Files" + sFilesAttached;
            }
            return Res;
        }

        /// <summary>
        /// Send a mail in HTML Format 
        /// </summary>
        /// <param name="Senders">List of senders for the mail</param>
        /// <param name="SendersCC">List of senders cc for the mail</param>
        /// <param name="SendersBCC">List of senders bcc for the mail</param>
        /// <param name="Subject">Subject for the mail</param>
        /// <param name="Body">Body of mail</param>
        /// <param name="ImagesAttached">List of images attached (only for HTML format)</param>
        /// <param name="FilesAttached">list of stream files attached to the mail</param>
        /// <param name="SSL">Is SSL required? </param>
        /// <param name="AuthenticationRequiered">The Authentication is required?</param>
        /// <param name="isHTML">The body of mail is HTML?</param>
        /// <param name="Result">Result of sending</param>
        /// <param name="mailPriority">Priority of mail</param>
        /// <returns>return true if the mail is submitted successfully</returns>
        public Boolean SendMail(List<String> Senders, List<String> SendersCC, List<String> SendersBCC, String Subject, String Body, List<ImageAttached> ImagesAttached, List<FileAttachedStream> FilesAttached, Boolean SSL, Boolean AuthenticationRequiered, Boolean isHTML, out String Result, System.Net.Mail.MailPriority mailPriority = System.Net.Mail.MailPriority.Normal)
        {
            Boolean Res = false;
            Result = String.Empty;
            String Sender = String.Empty;

            System.Net.Mail.MailMessage mailmsg = PreparateMailMessage(Senders, SendersCC, SendersBCC, Subject, Body, ImagesAttached, isHTML, out Sender);

            if (FilesAttached != null)
                FilesAttached.ForEach(fileattached => mailmsg.Attachments.Add(new System.Net.Mail.Attachment(fileattached.Content, fileattached.Name)));
            
            System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient();
            smtp.Host = Server;
            if (AuthenticationRequiered)
                smtp.Credentials = new System.Net.NetworkCredential(Mail, MailPassword);
            smtp.Port = Port;
            smtp.EnableSsl = SSL;
            try
            {
                smtp.Send(mailmsg);
                Result = "Submitted successfully";
                Res = true;
            }
            catch (Exception ex)
            {
                String sFilesAttached = String.Empty;
                if (FilesAttached != null)
                    FilesAttached.ForEach(fileattached => sFilesAttached += fileattached.Name + ", ");

                String Source = String.Empty;
                if (ex.Source != null)
                    Source = ex.Source;

                Result = "Error:" + ex.Message + " object:" + Source + " Remitente:" + Sender + " Files" + sFilesAttached;
            }
            return Res;
        }
        
        private static String SplitImages(String Contenido, out List<ImageAttached> imagenes)
        {
            imagenes = new List<ImageAttached>();
            int NumeroImg = 1;
            while (Contenido.IndexOf("\"data:image/") != -1)
            {
                ImageAttached imgAdj = new ImageAttached();
                String nImg = "000" + NumeroImg.ToString();
                int Posicion = Contenido.IndexOf("\"data:image/") + 1;
                int fin = Contenido.IndexOf("\"", Posicion);
                String DatosImgBase64 = Contenido.Substring(Posicion, fin - Posicion);
                String[] Partes1 = DatosImgBase64.Split(',');
                imgAdj.ImageBase64 = Partes1[1];
                imgAdj.Tipo = new System.Net.Mime.ContentType(Partes1[0].Replace("data:", "").Replace(";base64", ""));
                imgAdj.ContentId = nImg + "." + imgAdj.Tipo.MediaType.Replace("image/", "");
                imagenes.Add(imgAdj);
                Contenido = Contenido.Replace(DatosImgBase64, "cid:" + imgAdj.ContentId);
                NumeroImg++;
            }
            return Contenido;
        }

        private  System.Net.Mail.MailMessage PreparateMailMessage(List<String> Senders, List<String> SendersCC, List<String> SendersBCC, String Subject, String Body, List<ImageAttached> ImagesAttached,bool isHTML,out String Sender)
        {
            
            Sender = String.Empty;
            System.Net.Mail.MailMessage mailmsg = new System.Net.Mail.MailMessage();
            mailmsg.From = new System.Net.Mail.MailAddress(Mail);
            String SenderTemp = String.Empty;
            if (Senders != null)
            {
                Senders.ForEach(sender =>
                {
                    mailmsg.To.Add(sender);
                    SenderTemp += sender + ";";
                });
                Sender = SenderTemp;
            }
            else
            {
                throw new Exception("The sender can not be null");
            }

            if (SendersCC != null)
                SendersCC.ForEach(sendercc => mailmsg.CC.Add(sendercc));

            if (SendersBCC != null)
                SendersBCC.ForEach(senderbcc => mailmsg.Bcc.Add(senderbcc));

            if (isHTML)
            {
                List<ImageAttached> imagesTemp = null;

                if (Body.Contains("\"data:image/"))
                    Body = SplitImages(Body, out imagesTemp);

                if (ImagesAttached == null && imagesTemp != null)
                    ImagesAttached = new List<ImageAttached>();

                if (imagesTemp != null)
                    ImagesAttached.AddRange(imagesTemp);

                if (ImagesAttached != null)
                {
                    System.Net.Mail.AlternateView alternateView = System.Net.Mail.AlternateView.CreateAlternateViewFromString(Body, null, MediaTypeNames.Text.Html);
                    alternateView.ContentId = "htmlview";
                    alternateView.TransferEncoding = TransferEncoding.SevenBit;
                    ImagesAttached.ForEach(imageAttached =>
                    {
                        imageAttached.Tipo.Name = imageAttached.ContentId;
                        System.Net.Mail.LinkedResource linkedResource1 = new System.Net.Mail.LinkedResource(new System.IO.MemoryStream(System.Convert.FromBase64String(imageAttached.ImageBase64)), imageAttached.Tipo);
                        linkedResource1.ContentId = imageAttached.ContentId;
                        linkedResource1.TransferEncoding = TransferEncoding.Base64;
                        alternateView.LinkedResources.Add(linkedResource1);
                    });
                    mailmsg.AlternateViews.Add(alternateView);
                }
            }

            mailmsg.Subject = Subject;
            mailmsg.Body = Body;
            mailmsg.IsBodyHtml = isHTML;
            mailmsg.Priority = mailPriority;

            return mailmsg;
        }

    }

    public class FileAttachedStream
    {
        public FileAttachedStream()
        {

        }

        public FileAttachedStream(String name, System.IO.MemoryStream content)
        {
            Name = name;
            Content = content;
        }

        public String Name { set; get; }
        public System.IO.MemoryStream Content { set; get; }
        public System.IO.Stream StreamContent { set; get; }
    }

    public class ImageAttached
    {
        public String ImageBase64 { set; get; }
        public String ContentId { set; get; }
        public ContentType Tipo { set; get; }
    }

}
