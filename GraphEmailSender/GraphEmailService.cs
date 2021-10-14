using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace GraphEmailSender
{
    public class GraphEmailService
    {
        #region Data
        string clientId = "4cd1ad0e-ae0a-45c8-8282-c05b207006a8";
        string tenantId = "596cc758-b5ad-47aa-bdd7-89423f19a8a6";
        string clientSecret = "pZD7Q~VcmqfWYDklHfIQ3GVYJe9EAw3NvsfiH";
        string userId = "8d1a920a-63cd-4d9d-a9fc-358134363a76";

        

        string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

        string toEmailList = "krishanthadh@gmail.com,krishantha.dharmasiri@outlook.com,krishanthad@live.com";
        string subject = "NDB-2021-10-11-FD-New Investment-123";

        #endregion

        #region Main Functions
        
        /// <summary>
        /// Get the Access token from Azure AD 
        /// </summary>
        /// <returns></returns>
        public async Task<string> GetAuthTokenAsync()
        {
            IConfidentialClientApplication confidentialClient = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}/v2.0"))
                .Build();

            // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
            var authResult = await confidentialClient
                    .AcquireTokenForClient(scopes)
                    .ExecuteAsync().ConfigureAwait(false);

            return authResult.AccessToken;
        }

        /// <summary>
        /// Sends an Email
        /// </summary>
        /// <returns></returns>
        public async Task SendEmailAsync()
        {
            var token = await GetAuthTokenAsync();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );


            var message = GenerateEmailMessage();

            await graphServiceClient.Users[userId]
                  .SendMail(message, true)
                  .Request()
                  .PostAsync();            
        }

        /// <summary>
        /// Search the Inbox for the messages that contain the search phrase
        /// </summary>
        /// <param name="searchPhrase"></param>
        /// <returns></returns>
        public async Task<string> GetMessageIdBasedOnSubjectContentAsync(string searchPhrase)
        {
            var token = await GetAuthTokenAsync();

            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );

            //Create a custom query to search the messages based on the content in the subject
            List<QueryOption> queryOptions = new List<QueryOption>
            {
                new QueryOption("$search",searchPhrase)
            };

            var messages = await graphServiceClient.Users[userId].MailFolders["Inbox"].Messages.Request(queryOptions).GetAsync();

            //Get the ID of the very first message in the collection
            return messages[0].Id;            
        }

        /// <summary>
        /// Replys to a message, currently only reply to the email that the initial message came from
        /// AKA not a "Reply All"
        /// </summary>
        /// <returns></returns>
        public async Task ReplyToEmailAsync()
        {
            var token = await GetAuthTokenAsync();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );

            var mailId = await GetMessageIdBasedOnSubjectContentAsync("AAA");

            Message reply = graphServiceClient
                            .Users[userId]
                            .Messages[mailId]
                            .CreateReply()
                            .Request()
                            .PostAsync().Result;

            graphServiceClient
                .Users[userId]
                .Messages[reply.Id]
                .Request()
                .UpdateAsync(new Message()
                {
                    Body = new ItemBody()
                    {
                        Content = "<h2>This is the Reply</h2>",
                        ContentType = BodyType.Html
                    }
                })
                .Wait();


            graphServiceClient
                .Users[userId]
                .Messages[reply.Id]
                .Send()
                .Request()
                .PostAsync()
                .Wait();


        }

        #endregion


        #region Support Methods
        public async Task GetInboxMessgesAsync()
        {
            var token = await GetAuthTokenAsync();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );

            var messages = await graphServiceClient.Users[userId].MailFolders["Inbox"].Messages.Request().GetAsync();

            foreach (var m in messages)
            {

                Console.WriteLine($"The message with Subject : {m.Subject} recieved at : { m.ReceivedDateTime}");
            }
        }

        public async Task SearchEmailBySubjectAsync()
        {
            var token = await GetAuthTokenAsync();
            
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );

            //Create a custom query to search the messages using the subject line
            List<QueryOption> queryOptions = new List<QueryOption>
            {
                new QueryOption("$search","kkk")
            };

            var messages = await graphServiceClient.Users[userId].Messages.Request(queryOptions).GetAsync();

            Console.WriteLine($"Total Number, matched the search is : {messages.Count}");
            Console.WriteLine("---------------------------------------------------------------");
            foreach (var m in messages)
            {

                Console.WriteLine($"The message with Subject : {m.Subject} recieved at : { m.ReceivedDateTime}");
                Console.WriteLine("---------------------------------------------------------------");
            }
        }

        public async Task GetAllMessgesAsync()
        {
            var token = await GetAuthTokenAsync();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );

            var messages = await graphServiceClient.Users[userId].Messages.Request().GetAsync();

            foreach(var m in messages)
            {
               
                Console.WriteLine($"The message with Subject : {m.Subject} recieved at : { m.ReceivedDateTime}");
            }
        }

        public async Task GetMailFoldersAsync()
        {
            var token = await GetAuthTokenAsync();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {
                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("Bearer", token);
                })
                );

            var mailFolders = await graphServiceClient.Users[userId].MailFolders.Request().GetAsync();

            foreach (var m in mailFolders)
            {

                Console.WriteLine($"The Mailfolder Name : {m.DisplayName} and Id: { m.Id}");
            }
        }

        #endregion

        private Microsoft.Graph.Message GenerateEmailMessage()
        {
            // Build the email recipient list
            var emailsList = toEmailList.Split(',');
            var ToEmailList = new List<Recipient>();
            foreach (var e in emailsList)
            {
                var r = new Recipient();
                var a = new EmailAddress();
                a.Address = e;
                r.EmailAddress = a;

                ToEmailList.Add(r);
            }

            //Buiild the email message

            var message = new Microsoft.Graph.Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html,
                    Content = "<h2>The following deals creted by PM</h2><p>Deals table here....</p>"
                },
                ToRecipients = ToEmailList

            };

            return message;
        }
        
    }
}
