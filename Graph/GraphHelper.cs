using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;
using TimeZoneConverter;

namespace GraphTutorial
{
    public class GraphHelper
    {
        private static DeviceCodeCredential tokenCredential;
        private static GraphServiceClient graphClient;

        public static void Initialize(string clientId,
                                      string[] scopes,
                                      Func<DeviceCodeInfo, CancellationToken, Task> callBack)
        {
            tokenCredential = new DeviceCodeCredential(callBack, clientId);
            graphClient = new GraphServiceClient(tokenCredential, scopes);
        }

        public static async Task<string> GetAccessTokenAsync(string[] scopes)
        {
            var context = new TokenRequestContext(scopes);
            var response = await tokenCredential.GetTokenAsync(context);
            return response.Token;
        }

        public static async Task<IUserMessagesCollectionPage> GetMessagesAsync()
        {
            //var searchTerm = 'Undeliverable: UniSuper - Forgotten Username';

            var viewOptions = new List<QueryOption>
            {
                new QueryOption("$search","Forgotten"),
                //new QueryOption("endDateTime", endOfWeek.ToString("o"))
            };
            try
            {
                var messages = await graphClient
                               .Users["no-replydev@usmnpz.com.au"]
                               .Messages
                               .Request(viewOptions)
                               .Select("id,subject,body,attachments")
                               .GetAsync();

                return messages;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting messages user: {ex.Message}");
                return null;
            }
        }
    }
}