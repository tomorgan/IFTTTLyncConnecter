using System.Threading.Tasks;
using Microsoft.Rtc.Collaboration;
using System;

namespace LyncAsyncExtensionMethods
{
    public static class UserEndpointMethods
    {
        public static Task EstablishAsync(this 
            UserEndpoint endpoint)
        {
            Action<IAsyncResult> action = (result) =>
            {
                endpoint.EndEstablish(result);
            };

            return Task.Factory.FromAsync(
                endpoint.BeginEstablish,
                action,
                null);
        }

        public static Task TerminateAsync(this UserEndpoint endpoint)
        {
            return Task.Factory.FromAsync(

                endpoint.BeginTerminate, endpoint.EndTerminate, null);
        }
    }
}
