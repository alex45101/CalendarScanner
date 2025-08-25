using MimeKit;
using MailKit.Net.Imap;
using MailKit.Net.Smtp;
using MailKit;
using MailKit.Search;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace CalendarScanner
{
    public class MailSender
    {
        private readonly string mailServer, login, password;
        private readonly int port;
        private readonly bool ssl;

        /// <summary>
        /// Setting up the connection info
        /// </summary>
        /// <param name="mailServer">The smtp server</param>
        /// <param name="port">The port for the server</param>
        /// <param name="ssl">If the server is using ssl</param>
        /// <param name="login">The email for the account</param>
        /// <param name="password">The password for the given email</param>
        public MailSender(string mailServer, int port, bool ssl, string login, string password)
        {
            this.mailServer = mailServer;
            this.port = port;
            this.ssl = ssl;
            this.login = login;
            this.password = password;
        }

        /// <summary>
        /// Sending a message from the email. Automatically adds the from header for the message.
        /// </summary>
        /// <param name="message">The message to send</param>
        public void SendMessage(MimeMessage message)
        {
            //Add the from header automatically
            message.From.Add(new MailboxAddress("Work Schedular", login));

            using (var client = new SmtpClient())
            {
                //Connect to the smtp server
                client.Connect(mailServer, port, ssl);

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(login, password);

                client.Send(message);

                client.Disconnect(true);
            }
        }
    }

    /// <summary>
    /// Connection and accessing an email
    /// </summary>
    public class MailRepository
    {
        private readonly string mailServer, login, password;
        private readonly int port;
        private readonly bool ssl;

        /// <summary>
        /// Setting up the connection info
        /// </summary>
        /// <param name="mailServer">The imap4 server</param>
        /// <param name="port">The port for the server</param>
        /// <param name="ssl">If the server is using ssl</param>
        /// <param name="login">The email for the account</param>
        /// <param name="password">The password for the given email</param>
        public MailRepository(string mailServer, int port, bool ssl, string login, string password)
        {
            this.mailServer = mailServer;
            this.port = port;
            this.ssl = ssl;
            this.login = login;
            this.password = password;
        }

        /// <summary>
        /// Get all the unread emails
        /// </summary>
        /// <returns>The a collection of Mime Messages</returns>
        public IEnumerable<MimeMessage> GetUnreadMails()
        {
            var messages = new List<MimeMessage>();

            using (var client = new ImapClient())
            {
                client.Connect(mailServer, port, ssl);

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(login, password);

                // The Inbox folder is always available on all IMAP servers...
                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadWrite);
                var results = inbox.Search(SearchOptions.All, SearchQuery.Not(SearchQuery.Seen));
                foreach (var uniqueId in results.UniqueIds)
                {
                    var message = inbox.GetMessage(uniqueId);
                    messages.Add(message);

                    //Mark message as read
                    inbox.AddFlags(uniqueId, MessageFlags.Seen | MessageFlags.Flagged, true);
                }

                client.Disconnect(true);
            }

            return messages;
        }

        /// <summary>
        /// Get all the emails
        /// </summary>
        /// <returns>The collection of Mime Messages</returns>
        public IEnumerable<MimeMessage> GetAllMails()
        {
            var messages = new List<MimeMessage>();

            using (var client = new ImapClient())
            {
                client.Connect(mailServer, port, ssl);

                // Note: since we don't have an OAuth2 token, disable
                // the XOAUTH2 authentication mechanism.
                client.AuthenticationMechanisms.Remove("XOAUTH2");

                client.Authenticate(login, password);

                // The Inbox folder is always available on all IMAP servers...
                var inbox = client.Inbox;
                inbox.Open(FolderAccess.ReadOnly);
                var results = inbox.Search(SearchOptions.All, SearchQuery.All);
                foreach (var uniqueId in results.UniqueIds)
                {
                    var message = inbox.GetMessage(uniqueId);

                    messages.Add(message);
                }

                client.Disconnect(true);
            }

            return messages;
        }
    }
}
