using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;


namespace TimeTriggerAzureFun
{
    public static class AFAssignPermission
    {
        [FunctionName("AFAssignPermission")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            string responseMessage = string.IsNullOrEmpty(name)
                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
                : $"Hello, {name}. This HTTP triggered function executed successfully.";
            Microsoft.Graph.IDriveItemInviteCollectionPage result = null;
            try
            {
                string[] scopes = new string[] { $"https://graph.microsoft.com/.default" };
                // string[] scopes = new string[] { Environment.GetEnvironmentVariable("GraphScope") };
                // var credential = new Azure.Identity.ClientSecretCredential(Environment.GetEnvironmentVariable("AADTenantId"), Environment.GetEnvironmentVariable("AADClientId"), Environment.GetEnvironmentVariable("AADClientSecret")
                var credential = new Azure.Identity.ClientSecretCredential("973b3f01-ab0a-4620-b906-eb32095e50cc", "dd712a30-caac-45eb-b048-1cf1853b8dd8", "vSbPh-wRd-G4F865u58cD.Aj.ERqaUz-tP");
                // This example works with Microsoft.Graph 4+
                var httpClientnew = Microsoft.Graph.GraphClientFactory.Create(new Microsoft.Graph.TokenCredentialAuthProvider(credential, scopes));
                // Create a new instance of GraphServiceClient with the authentication provider.

                Microsoft.Graph.GraphServiceClient graphClient = new Microsoft.Graph.GraphServiceClient(httpClientnew);
                log.LogInformation($"Created graph client");

                var recipients = new System.Collections.Generic.List<Microsoft.Graph.DriveRecipient>()
                {
                    new Microsoft.Graph.DriveRecipient
                        {
                                Email = "harshal.desai@spadeworx.com"
                        }
                 };
                var message = "Testing - Here's the file that we're collaborating on. - (Mithun)";
                var requireSignIn = true;
                var sendInvitation = true;
                var roles = new System.Collections.Generic.List<String>()
                {
                "write"
                };

                result = await graphClient.Sites["76956606-5f0c-45e9-a4d0-7870674a3677"].Lists["7f036708-3d9a-46c9-a260-3b724b7e32d9"].Items["1"].DriveItem
                            .Invite(recipients, requireSignIn, roles, sendInvitation, message, null, null)
                            .Request()
                            .PostAsync();
            }
            catch (Exception ex)
            {
                log.LogInformation($"{ex.Message}");
            }
            return new OkObjectResult(responseMessage);
        }
    }
}
