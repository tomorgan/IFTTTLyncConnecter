using LyncAsyncExtensionMethods;
using Microsoft.Rtc.Collaboration;
using Microsoft.Rtc.Collaboration.Presence;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;

namespace EmailToIMClassLibrary
{
    public class LyncIM : IDisposable
    {
        UserEndpoint _endpoint { get; set; }
        Conversation _conversation { get; set; }
        CollaborationPlatform _collabPlatform { get; set; }
        string _sipaddress { get; set; }
        string _username { get; set; }
        string _password { get; set; }
        bool _endpointStarted = false;

        public async void Initialise(string sipAddress, string username, string password)
        {
            try
            {
                _sipaddress = sipAddress;
                _username = username;
                _password = password;
                Console.WriteLine("Establishing collab platform");
                await EstablishCollaborationPlatform();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public async void SendIMToLyncUser(List<string> messages, string toastMessage, string destinationSIP)
        {
            try
            {
                await Establish();
                await SendIM(messages, toastMessage, destinationSIP);
                await LogOut();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        private async Task EstablishCollaborationPlatform()
        {
            string userAgent = "IFTTT/Email Lync Instant Message Sender";
            var platformSettings = new ClientPlatformSettings(userAgent, Microsoft.Rtc.Signaling.SipTransportType.Tls);

            try
            {
                Console.WriteLine("Trying to start collab platform");
                _collabPlatform = new CollaborationPlatform(platformSettings);
                _collabPlatform.ProvisioningFailed += _collabPlatform_ProvisioningFailed;
                _collabPlatform.AllowedAuthenticationProtocol = Microsoft.Rtc.Signaling.SipAuthenticationProtocols.Ntlm;
                await _collabPlatform.StartupAsync();
                _endpointStarted = true;
                Console.WriteLine("Collab Platform Started.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error starting collab platform");
                Console.WriteLine(ex.ToString());
                if (ex.InnerException != null)
                {
                    Console.WriteLine(ex.InnerException.ToString());
                }
                else
                {
                    Console.WriteLine("Inner exception is empty");
                }
            }


            
        }
        
        private async Task Establish()
        {
            Console.WriteLine("Establishing with endpoint:" + _sipaddress);

            if (!_endpointStarted)
            {
                Console.WriteLine("Collab Platform not started, starting now");
                {
                   await EstablishCollaborationPlatform();
                }
            }
                     
            var settings = new UserEndpointSettings(_sipaddress);
            settings.Credential = new System.Net.NetworkCredential(_username, _password);

            
            _endpoint = new UserEndpoint(_collabPlatform, settings);
            await _endpoint.EstablishAsync();
        }

        void _collabPlatform_ProvisioningFailed(object sender, ProvisioningFailedEventArgs e)
        {
            Console.WriteLine("Provisioning failed: " + e.Exception.ToString());
        }

        private async Task SendIM(List<string> messages, string toastMessage, string destinationSip)
        {

            Conversation conversation = new Conversation(_endpoint);
            InstantMessagingCall call = new InstantMessagingCall(conversation);

            ToastMessage toast = new ToastMessage(toastMessage);
            CallEstablishOptions options = new CallEstablishOptions();

            await call.EstablishAsync(destinationSip, toast, options);
            foreach (var msg in messages)
            {
                await call.Flow.SendInstantMessageAsync(msg);
            }
            await conversation.TerminateAsync();
            Console.WriteLine("IM Sent");
        }

        private async Task LogOut()
        {
            await _endpoint.TerminateAsync();
        }

        public void Dispose()
        {
            _collabPlatform.ShutdownAsync();
        }
    }
}
