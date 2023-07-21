using System.Net;
using System.Net.Mail;
using System.Security;
using System.Text.RegularExpressions;

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
    /// </remarks>
    public static class MicroMailer
    {
        /// <summary>
        /// Haupt-Einstiegspunkt für die Anwendung.
        /// </summary>
        /// <param name="args">
        /// "Vishnu-Counter" "Vishnu-TreeParameters" "Vishnu-NodeId"
        /// "Betreff" "Meldungstext" "Mailserver:[Port]:[Passwort]" "Absender" "Empfäger mit Semikolon getrennt" ["Anhang" ...].
        /// Wichtig: die ersten drei Parameter werden im Vishnu-Betrieb von Vishnu generiert und werden in einer JobDescription.xml nicht aufgeführt;
        /// für den Test des MicroMailers müssn sie in den Debug-Parametern allerdings übergeben werden.
        /// </param>
        [STAThread]
        public static void Main(string[] args)
        {
            string paraMessage = EvaluateParametersOrDie(args, out string aufrufInfo, out int aufrufCounter, out string subject, out string messageText,
                out string host, out int port, out SecureString securePW, out string sender, out string recipients, out List<string> attachments);

            try
            {
                MicroMailer.quickSmtp(subject, messageText, host, port, securePW, sender, recipients, attachments);
            }
            catch (Exception ex)
            {
                SyntaxAndDie(ex.Message + Environment.NewLine + paraMessage);
            }

        }

        static void quickSmtp(string subject, string messageText, string host, int port, SecureString password,
            string sender, string recipients, List<string> attachments)
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

            using (SmtpClient smtpClient = port > 0 ? new SmtpClient(host, port) : new SmtpClient(host))
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
                foreach (string item in attachments)
                {
                    mailMessage.Attachments.Add(new Attachment(item));
                }
                smtpClient.Send(mailMessage);
                mailMessage.Dispose();
            }
        }

        private static string EvaluateParametersOrDie(string[] args, out string aufrufInfo, out int aufrufCounter, out string subject,
            out string messageText, out string host, out int port, out SecureString securePW, out string sender, out string recipients,
            out List<string> attachments)
        {
            // Übergabeparameter
            aufrufInfo = "";
            aufrufCounter = 0;
            subject = "";
            messageText = "";
            host = "";
            port = 0;
            securePW = new();
            sender = "";
            recipients = "";
            attachments = new();

            Exception? paraException = null;
            string paraMessage;

            string delim = "";
            try
            {
                _ = int.TryParse(args[0].Trim(), out aufrufCounter);
                args = args.Skip(3).ToArray();

                if (args.Length < 5) { SyntaxAndDie(null); }

                subject = args[0].TrimStart('"').TrimEnd('"');
                if (aufrufCounter < 0)
                {
                    aufrufInfo = "Das Problem ist behoben. Die ursprüngliche Meldung war:";
                    subject = String.Format("Das Problem ist behoben ({0}).", subject);
                    delim = Environment.NewLine;
                }
                messageText = aufrufInfo + delim + args[1].TrimStart('"').TrimEnd('"').Replace("#", Environment.NewLine);

                ParseHostPortPW(args[2].TrimStart('"').TrimEnd('"'), ref host, ref port, ref securePW);

                sender = args[3].TrimStart('"').TrimEnd('"');
                recipients = args[4].TrimStart('"').TrimEnd('"'); // Komma-separiert

                // mal schnell nen kleinen Aray-Slice ohne Extension-Methode:
                attachments
                  = new List<string>(Enumerable.Range(5, args.Length - 1).Where(n => n < args.Length).Select((n) => args[n].TrimStart('"').TrimEnd('"')));

            }
            catch (Exception ex)
            {
                paraException = ex;
            }
            finally
            {
                string attachmentsString = "";
                for (int i = 0; i < attachments.Count; i++)
                {
                    attachmentsString += attachments[i].ToString() + Environment.NewLine;
                }
                if (!String.IsNullOrEmpty(attachmentsString))
                {
                    attachmentsString = Regex.Replace(attachmentsString, ":.*", "");
                    attachmentsString = "\nAttachments: " + attachmentsString;
                }
                string portString = port == 0 ? "" : "\nPort: " + Regex.Replace(port.ToString() ?? "", ":.*?[ \n]", "\n");
                paraMessage = String.Format("{0}"
                  + "\nSubject: {1}"
                  + "\nMessage: {2}"
                  + "\nHost: {3}"
                  + portString
                  + "\nSender: {4}"
                  + "\nRecipients: {5}"
                  + attachmentsString
                  , Regex.Replace(paraException?.Message ?? "", ":.*?[ \n]", "\n")
                  , Regex.Replace(subject, ":.*", "")
                  , Regex.Replace(messageText, ":.*", "")
                  , Regex.Replace(host, ":.*", "")
                  , Regex.Replace(sender, ":.*", "")
                  , Regex.Replace(recipients, ":.*", ""));
            }
            if (paraException != null)
            {
                SyntaxAndDie(paraMessage);
            }
            return paraMessage;
        }

        private static void ParseHostPortPW(string hostPortPW, ref string host, ref int port, ref SecureString securePW)
        {
            try
            {
                Array.ForEach(hostPortPW.Split(':')[2].ToArray(), securePW.AppendChar);
            }
            catch { }
            securePW.MakeReadOnly();

            port = 0;
            try
            {
                if (Int32.TryParse(hostPortPW.Split(':')[1], out int paraPort))
                {
                    port = paraPort;
                };
            }
            catch { }

            host = hostPortPW.Split(':')[0];
        }

        private static void SyntaxAndDie(string? message)
        {
            string msg = message?.Trim() ?? "";
            if (!String.IsNullOrEmpty(msg))
            {
                msg += Environment.NewLine;
            }
            MessageBox.Show(msg
              + "Aufruf:\n"
              + "MicroMailer VishnuCounter Vishnu-TreeParameters Vishnu-NodeId Betreff Meldung Mailserver:[Port]:[Passwort]"
              + " \"Empfäger mit Semikolon getrennt\" [Anhang [...]]\n"
              + "Beispiel:\n"
              + "\"1\" \"Tree\" \"Node\" \"Test Fehler\" \"Dies ist eine Test-Message!\" \"smtp.<meinServer>.de::<Passwort>\""
              + " \"Vishnu@reallyhuman.net\" \"neteti@neteti.de;musterfrau@gmail.com\" \"Testdaten\\070605_DotNet_Gewuenschte_Abhaengigkeit.pdf\""
              + " \"Testdaten\\Parameterübergabe im Vishnu-Tree.pptx\"");
            Environment.Exit(255);
        }
    }
}
