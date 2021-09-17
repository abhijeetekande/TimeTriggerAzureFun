using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Azure.Cosmos.Table;
using System.Globalization;

namespace TimeTriggerAzureFun
{
    public static class ExpirationLinkEntry
    {
        [FunctionName("ExpirationLinkEntry")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            [Table("tblExpirationLinks", Connection = "AzureWebJobsStorage")] CloudTable cloudTable,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            try
            {
                log.LogInformation($"Processing request and making entry in Azure table");
                string test = data.expirationDate;
                string newdate = test.Substring(0, test.IndexOf("GMT")+3);

                DateTime expDate = DateTime.Parse(newdate, null, DateTimeStyles.AdjustToUniversal);
                log.LogInformation($"{String.Concat(Convert.ToString(data.targetUsers).Split('@')[0], Convert.ToString(data.ItemURL))}");
                TableOperation insertOrMergeOperation = TableOperation.InsertOrMerge(new ExpirationLinksTableEntity()
                {
                    PartitionKey = Guid.NewGuid().ToString(),// Convert.ToString(data.currentUser),  
                    RowKey = Guid.NewGuid().ToString(), //String.Concat(Convert.ToString(data.targetUsers).Split('@')[0], Convert.ToString(data.ItemURL)),// Guid.NewGuid().ToString(),//$"{Convert.ToString(data.targetUsers).Split('@')[0]}{data.ItemURL}",
                    ItemURL = data.ItemURL,
                    SharedByUser = data.currentUser,
                    SharedWithUser = data.targetUsers,
                    PermissionLevel = Convert.ToBoolean(data.editingEnabled) ? "Edit" : "Read",
                    ExpirationDate = expDate.ToShortDateString(),//ToString("dd/mm/yyyy"), //data.expirationDate.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK"),
                    WebURL = data.webUrl,
                    Expired = Convert.ToBoolean(false),
                    SiteID = data.SiteID,
                    ListID = data.ListID,
                    ItemID = data.ItemID,
                    PermissionID = data.PermissionID

                }) ;
                TableResult result = cloudTable.Execute(insertOrMergeOperation);
                       
            }
            catch (Exception ex)
            {
                log.LogInformation($"Error: {ex.StackTrace}");
            }

            string responseMessage = $"Hello, {data.currentUser}. your shared link entered in azure table.";
            log.LogInformation($"Azure function executed succesfully!!");

            return new OkObjectResult(responseMessage);
        }

        public class ExpirationLinksTableEntity : TableEntity
        {
            public string ItemURL { get; set; }
            public string SharedByUser { get; set; }
            public string SharedWithUser { get; set; }
            public string PermissionLevel { get; set; }
            public string ExpirationDate { get; set; }
            public string WebURL { get; set; }
            public Boolean Expired { get; set; }
            public string SiteID { get; set; }
            public string ListID { get; set; }
            public string ItemID { get; set; }
            public string PermissionID { get; set; }

        }
    }
}
