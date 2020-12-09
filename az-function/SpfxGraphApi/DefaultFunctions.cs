using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using System.Net.Http;
using System.Net.Http.Headers;
using System;

namespace SpfxGraphApi
{
    public class DefaultFunctions
    {
        private AppInfo _appInfo;

        public DefaultFunctions(AppInfo appInfo)
        {
            _appInfo = appInfo;
        }

        [FunctionName("SaveMail")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "SaveMail/{tenantId}/{mailId}")] HttpRequestMessage request, string tenantId, string mailId,
            ILogger log)
        {
            var accessToken = request.Headers.Authorization.Parameter;
            var graphClient = CreateGraphClient(tenantId, accessToken);
            var mail = await graphClient.Me.Messages[mailId].Request().GetAsync();
            var mailStream = await graphClient.Me.Messages[mailId].Content.Request().GetAsync();

            // upload to root OneDrive folder
            await graphClient.Me.Drive.Root.ItemWithPath(mail.Subject + ".eml").Content.Request().PutAsync<DriveItem>(mailStream);

            return new OkResult();
        }

        public GraphServiceClient CreateGraphClient(string tenantId, string accessToken)
        {
            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                        .Create(_appInfo.ClientId)
                        .WithClientSecret(_appInfo.ClientSecret)
                        .WithTenantId(tenantId)
                        .Build();
            var userAssertion = new UserAssertion(accessToken);
            var authProvider = new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        try
                        {
                            var tokenResult = await confidentialClientApplication.AcquireTokenOnBehalfOf(new string[] { "https://graph.microsoft.com/.default" }, userAssertion).ExecuteAsync().ConfigureAwait(false);

                            requestMessage.Headers.Authorization =
                                new AuthenticationHeaderValue("Bearer", tokenResult.AccessToken);
                        }
                        catch(Exception ex)
                        {

                        }
                    });

            return new GraphServiceClient(authProvider);
        }
    }
}
