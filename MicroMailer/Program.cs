using NetEti.ApplicationEnvironment;
using System.Net;
using System.Net.Mail;
using System.Security;
using System.Text;

namespace NetEti.Communication
{
    /// <summary>
    /// Exe für den E-Mail-Versand.
    /// Achtung: Demo-Programm, muss für produktiven Gebrauch angepasst werden;
    /// Mail-Server, Port und Passwort werden hier im Klartext übergeben.
    /// </summary>
    /// <remarks>
    /// File: MicroMailer.cs
    /// Autor: Erik Nagel, NetEti
    ///
    /// 21.06.2015 Erik Nagel, NetEti: erstellt
    /// 21.04.2023 Erik Nagel: für .Net 7.0 komplett überarbeitet.
    /// 10.08.2023 Erik Nagel: Auf benannte Parameter umgestellt.
    ///                        Dadurch wird eine erweiterte Parameterersetzung ermöglicht.
    /// </remarks>
    public static class MicroMailer
    {
        /// <summary>
        /// Haupt-Einstiegspunkt für die Anwendung.
        /// </summary>
        /// <param name="args">
        /// "Vishnu-Counter" "Vishnu-TreeParameters" "Vishnu-NodeId"
        /// "Betreff" "Meldungstext" "Mailserver:[Port]:[Passwort]" "Absender" "Empfänger mit Semikolon getrennt" ["Anhang" ...].
        /// Wichtig: die ersten drei Parameter werden im Vishnu-Betrieb von Vishnu generiert und werden in einer JobDescription.xml nicht aufgeführt;
        /// für den Test des MicroMailers müssen sie in den Debug-Parametern allerdings übergeben werden.
        /// </param>
        [STAThread]
        public static void Main(string[] args)
        {
            string paraMessage = EvaluateParametersOrDie(out string subject, out string messageText,
                out string host, out int? port, out SecureString? securePW, out string sender,
                out string recipients, out List<string> attachments);

            try
            {
                MicroMailer.QuickSmtp(subject, messageText, host, port, securePW, sender, recipients, attachments);
            }
            catch (Exception ex)
            {
                Die<string>(ex.Message, paraMessage);
            }
        }

        static void QuickSmtp(string subject, string messageText, string host, int? port, SecureString? password,
            string sender, string recipients, List<string>? attachments)
        {
            /*
             * Funktioniert so nicht, da ein Dialog aufpoppt; theoretisch denkbar: VB-Script mit Sleep und SendKey oder AutoIt-Steuerung.
             * Bleibt aber eine unsaubere Lösung:
             * 
             * string mailto = string.Format("mailto:{0}?Subject={1}&Body={2}", recipients, subject, messageText);
             * mailto = Uri.EscapeUriString(mailto);
             * System.Diagnostics.Process.Start(new ProcessStartInfo(mailto) { UseShellExecute = true });
             * return;
            */

            using (SmtpClient smtpClient = port != null ? new SmtpClient(host, (int)port) : new SmtpClient(host))
            {
                smtpClient.Credentials = new NetworkCredential(sender, password);
                smtpClient.EnableSsl = true;
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(sender);
                foreach (string recipient in recipients.Split(';'))
                {
                    mailMessage.To.Add(recipient);
                }
                mailMessage.Subject = subject;
                mailMessage.Body = messageText;
                if (attachments != null)
                {
                    foreach (string item in attachments)
                    {
                        mailMessage.Attachments.Add(new Attachment(item));
                    }
                }
                /*
                int l = password?.Length ?? 0;
                string toString = String.Join(';', mailMessage.To);
                MessageBox.Show(String.Format($"host: {host}, pwLength: {l}, port: {port}")
                    + Environment.NewLine + String.Format($"sender: {sender}")
                    + Environment.NewLine + String.Format($"to: {toString}")
                    + Environment.NewLine + String.Format($"subject: {mailMessage.Subject}")
                    + Environment.NewLine + String.Format($"body: {mailMessage.Body}"),
                    "Debug-Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                */
                
                smtpClient.Send(mailMessage);
                mailMessage.Dispose();
            }
        }

        private static string EvaluateParametersOrDie(out string caption, out string message,
            out string host, out int? port, out SecureString? securePW,
            out string sender, out string recipients,
            out List<string> attachments)
        {
            attachments = new();
            CommandLineAccess commandLineAccess = new();

            string? tmpStr = commandLineAccess.GetStringValue("EscalationCounter", "0");
            bool isResetting = Int32.TryParse(tmpStr, out int escalationCounter) && escalationCounter < 0;
            caption = commandLineAccess.GetStringValue("Caption", "Information") ?? "Information";

            string msg = commandLineAccess.GetStringValue("Message", null)
                ?? Die<string>("Es muss ein Meldungstext mitgegeben werden.", commandLineAccess.CommandLine);

            tmpStr = commandLineAccess.GetStringValue("MessageNewLine", null);
            string[] messageLines;
            if (!String.IsNullOrEmpty(tmpStr))
            {
                messageLines = msg.Split(tmpStr);
            }
            else
            {
                messageLines = new string[1] { msg };
            }
            string delim = "";
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < messageLines.Count(); i++)
            {
                sb.Append(delim + messageLines[i]);
                delim = Environment.NewLine;
            }
            message = sb.ToString();
            if (isResetting)
            {
                string resolvedPrefix = commandLineAccess.GetStringValue(
                    "ResolvedPrefix", "Das Problem ist behoben. Die ursprüngliche Meldung war:") ?? "";
                message = resolvedPrefix
                        + Environment.NewLine
                + message;
                caption = "(" + caption + ")";
            }

            string mailHostPort = commandLineAccess.GetStringValue("MailHostPort", null)
                ?? Die<string>("Es muss Mail-Host mitgegeben werden (optional mit Port nach Doppelpunkt).",
                commandLineAccess.CommandLine);

            string? mailPassword = commandLineAccess.GetStringValue("MailPassword", null);
            ParseHostPortPW(mailHostPort, mailPassword, out host, out port, out securePW);

            sender = commandLineAccess.GetStringValue("MailSender", null)?.TrimStart('"').TrimEnd('"')
                ?? Die<string>("Es muss ein Absender mitgegeben werden.", commandLineAccess.CommandLine);

            recipients = commandLineAccess.GetStringValue("MailRecipients", null)?.TrimStart('"').TrimEnd('"')
                ?? Die<string>("Es müssen ein oder mehrere Empfänger mitgegeben werden (Trennzeichen ist Semikolon).",
                commandLineAccess.CommandLine);

            string? attachmentsString
                = commandLineAccess.GetStringValue("MailAttachments", null)?.TrimStart('"').TrimEnd('"');
            if (!String.IsNullOrEmpty(attachmentsString))
            {
                attachments = new List<string>(attachmentsString.Split(';'));
            }

            string paraMessage = String.Format(
                    "\nCaption: {0}"
                  + "\nMessage: {1}"
                  + "\nHostPort: {2}"
                  + "\nSender: {3}"
                  + "\nRecipients: {4}\n"
                  + attachmentsString
                  , caption
                  , message
                  , mailHostPort
                  , sender
                  , recipients
            );
            return paraMessage;
        }

        private static void ParseHostPortPW(string mailHostPort, string? mailPassword, out string host,
            out int? port, out SecureString? securePW)
        {
            string[] mailHostPortArray = mailHostPort.TrimStart('"').TrimEnd('"').Split(':');
            host = mailHostPortArray[0];
            port = null;
            if (mailHostPortArray.Length > 1)
            {
                if (Int32.TryParse(mailHostPortArray[1], out int paraPort))
                {
                    port = paraPort;
                }
            }
            securePW = null;
            if (!String.IsNullOrEmpty(mailPassword))
            {
                securePW = new();
                Array.ForEach(mailPassword.ToArray(), securePW.AppendChar);
                securePW.MakeReadOnly();
            }
        }

        private static T Die<T>(string? message, string? commandLine = null)
        {
            string usage = "Syntax:"
                + Environment.NewLine
                + "\t-Message=<Nachricht>"
                + Environment.NewLine
                + "\t-MailHostPort=<Host[:Port]>"
                + Environment.NewLine
                + "\t-MailSender=<Absender>"
                + Environment.NewLine
                + "\t-MailRecipients=<Empfänger[;Empfänger...]>"
                + Environment.NewLine
                + "\t[-MailPassword=<Passwort>]"
                + Environment.NewLine
                + "\t[-Caption=<Überschrift>]"
                + Environment.NewLine
                + "\t[-MessageNewLine=<NewLine-Kennung>]"
                + Environment.NewLine
                + "\t[-MailArrachments=<Mail-Anhang[;Mail-Anhang...]]>"
                + Environment.NewLine
                + "\t[-EscalationCounter={-n;+n} (negativ: Ursache behoben)]"
                + Environment.NewLine
                + "\t[-ResolvedPrefix=<Vorangestellter Kurztext bei negativem EscalationCounter>]"
                + Environment.NewLine
                + "Beispiel:"
                + Environment.NewLine
                + "\t-Message=\"Server-1:#Zugriffsproblem#Connection-Error...\""
                + Environment.NewLine
                + "\t-Caption=\"SQL-Exception\""
                + Environment.NewLine
                + "\t-ResolvedPrefix=\"No longer valid:\""
                + Environment.NewLine
                + "\t-MessageNewLine=\"#\""
                + Environment.NewLine
                + "\t-MailHostPort=\"MyHost:587\""
                + Environment.NewLine
                + "\t-MailSender=\"me@mymail\""
                + Environment.NewLine
                + "\t-MailRecipients=\"a@amail;b@bmail\""
                + Environment.NewLine
                + "\t-MailAttachments=\"a-file_path;another-file-path\"";
            if (commandLine != null)
            {
                usage = "Kommandozeile: " + commandLine + Environment.NewLine + usage;
            }
            MessageBox.Show(message + Environment.NewLine + usage, "Mail Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            throw new ArgumentException(message + Environment.NewLine + usage);
        }
    }
}
