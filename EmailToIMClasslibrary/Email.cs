using ImapX;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EmailToIMClassLibrary
{
    public class NewEmailEventArgs : EventArgs
    {
        public NewEmailEventArgs(List<string> messages)
        {
            Messages = messages;
        }
        public List<string> Messages { get; set; }
    }

    public class Email
    {
        public event EventHandler<NewEmailEventArgs> NewMessages = delegate { };
        private ImapClient _client;

        private string _IMAPClientUri;
        private string _IMAPUsername;
        private string _IMAPPassword;
        private string _subjectKeyword;

        public Email(string imapClientURI, string imapClientUsername, string imapClientPassword, string subjectKeyword)
        {
            _IMAPClientUri = imapClientURI;
            _IMAPUsername = imapClientUsername;
            _IMAPPassword = imapClientPassword;
            _subjectKeyword = subjectKeyword;
        }

        public void CheckNow()
        {
            var newMessages = CheckForMessages();
            if (newMessages.Any())
            {
                NewMessages(this, new NewEmailEventArgs(newMessages));
            } 
        }
             
               
        private List<string> CheckForMessages()
        {
            var messages = new List<string>();

            _client = new ImapClient(_IMAPClientUri, true);
            if (_client.Connect() && _client.Login(_IMAPUsername, _IMAPPassword))
            {
                try
                {
                    var keyword = _subjectKeyword;                    
                    var emails = _client.Folders.Inbox.Search(string.Format("UNSEEN SUBJECT \"{0}\"", keyword), ImapX.Enums.MessageFetchMode.Full);
                    Console.WriteLine(string.Format("{0} emails", emails.Count()));
                    foreach (var email in emails)
                    {
                        messages.Add(email.Body.Text);
                        email.Remove();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    _client.Logout();
                    _client.Disconnect();                    
                }
            }
            else
            {
                Console.WriteLine("Bad email login");
            }
            return messages;
        }
    }
}
