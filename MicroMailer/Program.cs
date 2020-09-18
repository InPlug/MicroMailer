using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Net.Mail;
using System.Net;

namespace NetEti.Communication
{
    /// <summary>
    /// Exe für den E-Mail-Versand.
    /// Achtung: Demo-Programm, muss u.U. für produktiven Gebrauch angepasst werden,
    /// da Mail-Server und Passwort im Klartext übergeben werden.
    /// </summary>
    /// <remarks>
    /// File: MicroMailer.cs
    /// Autor: Erik Nagel, NetEti
    ///
    /// 21.06.2015 Erik Nagel, NetEti: erstellt
    /// </remarks>
    public static class MicroMailer
    {
        static private string subject;
        static private string messageText;
        static private string host;
        static private string sender;
        static private string recipients;
        static private List<string> attachments;

        /// <summary>
        /// Haupt-Einstiegspunkt für die Anwendung.
        /// </summary>
        /// <param name="args">Betreff|Meldung|Mailserver[:Port[:Passwort]]|Absender|Empfäger mit Semikolon getrennt.</param>
        [STAThread]
        public static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string delim = "";
            string aufrufInfo = "Der Aufruf-Zähler fehlt!";
            string countString = "";
            int aufrufCounter = 0;
            try
            {
                countString = args[0];
                Int32.TryParse(countString, out aufrufCounter);
                args = args.Skip(3).ToArray();
                subject = args[0].TrimStart('"').TrimEnd('"');
                if (aufrufCounter < 0)
                {
                    aufrufInfo = "Das Problem ist behoben. Die ursprüngliche Meldung war:";
                    subject = String.Format("Das Problem ist behoben ({0}).", subject);
                    delim = Environment.NewLine;
                }
                else
                {
                    //aufrufInfo = countString + ". Warnung";
                    aufrufInfo = "";
                }
                messageText = aufrufInfo + delim + args[1].TrimStart('"').TrimEnd('"').Replace("#", Environment.NewLine);
                host = args[2].TrimStart('"').TrimEnd('"');
                sender = args[3].TrimStart('"').TrimEnd('"');
                recipients = args[4].TrimStart('"').TrimEnd('"'); // komma-separiert
                                                                  // mal schnell nen kleinen Aray-Slice ohne Extension-Methode:
                attachments
                  = new List<string>(Enumerable.Range(5, args.Length - 1).Where(n => n < args.Length).Select((n) => args[n].TrimStart('"').TrimEnd('"')));
                MicroMailer.quickSmtp(subject, messageText, host, sender, recipients, attachments);
            }
            catch // (Exception ex)
            {
                MessageBox.Show(String.Format("MicroMailer: {0}!"
                  + "\nSubject: {1}"
                  + "\nMessage: {2}"
                  + "\nHost: {3}"
                  + "\nSender: {4}"
                  + "\nRecipients: {5}"
                  //+ "\nAttachments: {6}"
                  //, ex.Message
                  , "Fehler beim Mailversand"
                  , subject ?? "><"
                  , messageText ?? "><"
                  , host ?? "><"
                  , sender ?? "><"
                  , recipients ?? "><"
                  , subject ?? "><"));
            }
        }

        static void quickSmtp(string subject, string messageText, string hostPort, string sender, string recipients, List<string> attachments)
        {
            String[] hostNPortNPassword = hostPort.Split(':');
            string host = hostNPortNPassword[0];
            string password = "";
            int port = 25;
            int paraPort = -1;
            if (hostNPortNPassword.Length > 1)
            {
                Int32.TryParse(hostNPortNPassword[1], out paraPort);
                if (paraPort > 0)
                {
                    port = paraPort;
                }
            }
            if (hostNPortNPassword.Length > 2)
            {
                password = hostNPortNPassword[2];
            }
            using (SmtpClient smtpClient = new SmtpClient(host, port))
            {
                smtpClient.Credentials = new NetworkCredential(sender, password);
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(sender);
                mailMessage.To.Add(recipients);
                mailMessage.Subject = subject;
                mailMessage.Body = messageText;
                foreach (string item in attachments)
                {
                    mailMessage.Attachments.Add(new Attachment(item));
                }
                smtpClient.Send(mailMessage);
                mailMessage.Dispose();
            }
        }
    }
}
