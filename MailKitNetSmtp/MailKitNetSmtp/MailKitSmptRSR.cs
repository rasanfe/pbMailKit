using System;
using System.IO;
using System.Collections.Generic;
using MailKit.Net.Smtp;
using MimeKit;
using MailKit.Security;

namespace MailKitNetSmtp
{
    public class MailKitSmptRSR
    {
        internal string errorText = "";
        internal string senderMail = "", recipientMail = "";
        internal string smtpMessage = "", smtpSubject = "", smtpServer = "", smtpUserName = "", smtpPassword = "", mailerName = "", smtpCharSet = "";
        internal List<string> attachments;
        internal bool readreceipt, html = false, useSSL = true;
        internal int smtpPort = 465, smtpConnecttype = 1, smtpAuthmethod = 5;


        // Connection Type
        internal const int None = 0;
        internal const int Auto = 1;
        internal const int SslOnConnect = 2;
        internal const int StartTls = 3;
        internal const int StartTlsWhenAvailable = 4;

        // Priority
        internal const int NoPriority = 0;
        internal const int LowPriority = 1;
        internal const int NormalPriority = 2;
        internal const int HighPriority = 3;

        SmtpClient smtp;
        MimeMessage mail;


        //Constructor
        public MailKitSmptRSR()
        {
            attachments = new List<string>();
            smtp = new SmtpClient();
            mail = new MimeMessage();
        }

        internal MailboxAddress emailName(string name, string mail)
        {
            if (String.IsNullOrEmpty(name))
            {
                return new MailboxAddress(mail, mail);
            }
            else
            {
                return new MailboxAddress(name, mail);
            }
        }
        public int Send()
        {
            int result = 0;
            result = SmtpConnect();
            if (result == -1) { return result; }
            result = SmtpSend();
            if (result == -1) { return result; }
            result = SmtpDisconnect();
            return result;
        }
        public void SetMessage(string pbmessage)
        {
            SetMessage(pbmessage, false);
        }
        public void SetMessage(string pbmessage, bool pbHTML)
        {
            smtpMessage = pbmessage;
            html = pbHTML;
        }
        public void SetRecipientEmail(string pbRecipientName, string pbRecipientMail)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbRecipientMail))
            {
                errorText = "To Mail cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbRecipientMail), message: errorText);
            }
            mail.To.Add(emailName(pbRecipientName, pbRecipientMail));
            recipientMail = pbRecipientMail;
        }
        public void SetCCRecipientEmail(string pbCCrecipientName, string pbCCrecipientMail)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbCCrecipientMail))
            {
                errorText = "CC Mail cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbCCrecipientMail), message: errorText);
            }
            mail.Cc.Add(emailName(pbCCrecipientName, pbCCrecipientMail));
        }
        public void SetBCCRecipientEmail(string pbBCCrecipientName, string pbBCCrecipientMail)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbBCCrecipientMail))
            {
                errorText = "Bcc Mail cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbBCCrecipientMail), message: errorText);
            }
            mail.Bcc.Add(emailName(pbBCCrecipientName, pbBCCrecipientMail));
        }
        public void SetReplyToEmail(string pbReplytoName, string pbReplytoMail)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbReplytoMail))
            {
                errorText = "ReplyTo Mail cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbReplytoMail), message: errorText);
            }

            mail.ReplyTo.Add(emailName(pbReplytoName, pbReplytoMail));
        }
        public void SetSenderEmail(string pbSenderName, string pbSenderMail)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbSenderMail))
            {
                errorText = "From Mail cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbSenderMail), message: errorText);
            }
            mail.From.Add(emailName(pbSenderName, pbSenderMail));
            senderMail = pbSenderMail;
        }
        public void SetSMTPServer(string pbsmtpserver)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbsmtpserver))
            {
                errorText = "Server cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbsmtpserver), message: errorText);
            }
            smtpServer = pbsmtpserver;
        }
        public void SetSubject(string pbsubject)
        {
            smtpSubject = pbsubject;
        }
        public void SetAttachment(string pbattachment)
        {
            attachments.Add(@pbattachment);
        }
        public void SetCharSet(string pbcharset)
        {
            smtpCharSet = pbcharset;
        }
        public void SetUsernamePassword(string pbusername, string pbpassword)
        {
            ResetErrorMessage();
            if (String.IsNullOrEmpty(pbusername))
            {
                errorText = "User Name cannot be null";
                throw new ArgumentNullException(paramName: nameof(pbusername), message: errorText);
            }

            smtpUserName = pbusername;
            smtpPassword = pbpassword;
        }
        public void SetPort(int pbport)
        {
            smtpPort = pbport;
        }

        public void SetAuthMethod(int pbauthmethod)
        {
            smtpAuthmethod = pbauthmethod;
        }
        public void SetConnectionType(int pbconnecttype)
        {
            smtpConnecttype = pbconnecttype;

        }
        public string GetLastErrorMessage()
        {
            return errorText;
        }
        internal void ResetErrorMessage()
        {
            errorText = "";
        }
        public void SetMailerName(string pbmailername)
        {
            mailerName = pbmailername;
        }
        public void SetPriority(int pbpriority)
        {
            switch (pbpriority)
            {
                case LowPriority:
                    SetPriorityLow();
                    break;
                case NormalPriority:
                    SetPriorityNormal();
                    break;
                case HighPriority:
                    SetPriorityHigh();
                    break;
                default:
                    SetPriorityNone();
                    break;
            }

        }
        public void SetPriorityNone() { }
        public void SetPriorityLow()
        {
            mail.Priority = MessagePriority.NonUrgent;

        }
        public void SetPriorityNormal()
        {
            mail.Priority = MessagePriority.Normal;
        }
        public void SetPriorityHigh()
        {
            mail.Priority = MessagePriority.Urgent;
        }
        public void SetReadReceiptRequested(bool pbreadreceipt)
        {
            readreceipt = pbreadreceipt;
        }
        public int SmtpConnect()
        {
            ResetErrorMessage();
            try
            {

                //Read Requested
                if (readreceipt == true) { mail.Headers[HeaderId.DispositionNotificationTo] = senderMail; }
                //XMailer
                if (mailerName != "") { mail.Headers[HeaderId.XMailer] = mailerName; }


                switch (smtpConnecttype)
                {
                    case Auto:
                        smtp.Connect(smtpServer, smtpPort, SecureSocketOptions.Auto);
                        break;
                    case SslOnConnect:
                        smtp.Connect(smtpServer, smtpPort, SecureSocketOptions.SslOnConnect);
                        break;
                    case StartTls:
                        smtp.Connect(smtpServer, smtpPort, SecureSocketOptions.StartTls);
                        break;
                    case StartTlsWhenAvailable:
                        smtp.Connect(smtpServer, smtpPort, SecureSocketOptions.StartTlsWhenAvailable);
                        break;
                    default:
                        smtp.Connect(smtpServer, smtpPort, SecureSocketOptions.None);
                        break;
                }

                smtp.Authenticate(smtpUserName, smtpPassword);
                return 1;
            }
            catch (Exception ex)
            {
                errorText = "Connect Error " + ex.Message;
                return -1;
            }

        }
        public int SmtpSend()
        {
            ResetErrorMessage();
            try
            {
                if (String.IsNullOrEmpty(senderMail) || String.IsNullOrEmpty(recipientMail))
                {
                    errorText = "The sender and the recipient is required";
                    return -1;
                }

                //Asunto
                mail.Subject = smtpSubject;


                //Cuerpo del Mensaje

                var builder = new BodyBuilder();

                if (html == true)
                {
                    builder.HtmlBody = smtpMessage;
                }
                else
                {
                    builder.TextBody = smtpMessage;
                }


                //Adjuntos
                foreach (var attachment in attachments)
                {
                    builder.Attachments.Add(attachment);
                }

                mail.Body = builder.ToMessageBody();


                //Envio el Maail
                smtp.Send(mail);
                return 1;
            }
            catch (Exception ex)
            {
                errorText = "Send Error " + ex.Message;
                return -1;
            }

        }
        public int SmtpDisconnect()
        {
            ResetErrorMessage();
            try
            {
                smtp.Disconnect(true);
                return 1;
            }
            catch (Exception ex)
            {
                errorText = "Disconnect Error " + ex.Message;
                return -1;
            }

        }
    }
}
