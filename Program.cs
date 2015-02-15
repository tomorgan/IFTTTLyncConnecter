using Microsoft.Rtc.Collaboration;
using System;
using System.Linq;
using LyncAsyncExtensionMethods;
using System.Threading.Tasks;
using EmailToIMClassLibrary;
using System.Threading;


namespace EmailToIM
{
    class Program
    {

        private static EmailToIMClassLibrary.Email emailChecker;
        private static EmailToIMClassLibrary.LyncIM lyncIM;

        static string imapclient { get; set; }
        static string imapusername { get; set; }
        static string imappassword { get; set; }
        static string subjectKeyword { get; set; }

        static string destinationSIP { get; set; }
        static string sipAddress { get; set; }
        static string sipUsername { get; set; }
        static string sipPassword { get; set; }
        static string IMToastMsg { get; set; }


        static void Main(string[] args)
        {
            try
            {
                GetConfigValues();
                emailChecker = new EmailToIMClassLibrary.Email(imapclient, imapusername, imappassword, subjectKeyword);
                emailChecker.NewMessages += emailChecker_NewMessages;

                Console.WriteLine("Initialising lync im");
                lyncIM = new EmailToIMClassLibrary.LyncIM();
                lyncIM.Initialise(sipAddress, sipUsername, sipPassword);

                while (true)
                {
                    emailChecker.CheckNow();
                    Thread.Sleep(30000);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        static void emailChecker_NewMessages(object sender, EmailToIMClassLibrary.NewEmailEventArgs e)
        {
            Console.WriteLine("emailChecker_NewMessages");
            lyncIM.SendIMToLyncUser(e.Messages, IMToastMsg, destinationSIP);
        }

        private static void GetConfigValues()
        {
            var settings = EmailToIM.Properties.Settings.Default;
            imapclient = settings.emailIMAPClient;

            imapusername = settings.emailUsername;

            imappassword = settings.emailPassword;
            subjectKeyword = settings.emailSubjectKeyword;
            
            sipAddress = settings.sipAddress;
            sipUsername = settings.username;
            sipPassword = settings.password;


            destinationSIP = settings.destinationSIP;
            
            IMToastMsg = settings.IMToastMsg;                       
        }


    }
}
