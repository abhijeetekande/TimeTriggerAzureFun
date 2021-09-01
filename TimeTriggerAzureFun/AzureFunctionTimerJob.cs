using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Microsoft.Extensions.Logging;
using Microsoft.Azure.Cosmos.Table;
using Microsoft.Graph;
using Microsoft.Azure;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;

namespace TimeTriggerAzureFun
{
    public static class AzureFunctionTimerJob
    {
        private static HttpClient _client = new HttpClient();
        [FunctionName("AzureFunctionTimerJob")]
        public static async Task Run([TimerTrigger("0 */5 * * * *")] TimerInfo myTimer, [Table("tblExpirationLinks", Connection = "AzureWebJobsStorage")] CloudTable cloudTable, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
            try
            {
                DateTime accMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                string formatted = accMonth.ToLongTimeString();
                string permissionid = "";
                TableQuery<ExpirationLinksTableEntity> qEmptyPermId = new TableQuery<ExpirationLinksTableEntity>()
                    .Where(TableQuery.GenerateFilterCondition("PermissionID", QueryComparisons.Equal, ""));
                foreach (ExpirationLinksTableEntity entity in cloudTable.ExecuteQuery(qEmptyPermId))
                {
                    permissionid = await GetUnresolvedUserPermId(entity, log);
                    UpdatePermissionIDInTable(entity, cloudTable, permissionid, log);
                }

                    TableQuery<ExpirationLinksTableEntity> query = new TableQuery<ExpirationLinksTableEntity>()
                    .Where(TableQuery.CombineFilters(
                    TableQuery.GenerateFilterCondition("ExpirationDate", QueryComparisons.Equal, DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK")),
                    TableOperators.And,
                    TableQuery.GenerateFilterConditionForBool("Expired", QueryComparisons.Equal, false))); //DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK")
                foreach (ExpirationLinksTableEntity entity in cloudTable.ExecuteQuery(query))
                {

                    //var permissionid = await GetUnresolvedUserPermId(entity, log);
                    // var entities = new IEnumerable<ExpirationLinksTableEntity>();
                    //if (entity.PermissionID == "")
                    //{
                    //    permissionid = await GetUnresolvedUserPermId(entity, log);
                    //    UpdatePermissionIDInTable(entity, cloudTable, permissionid, log);
                    //}
                    //else
                    //{
                    //    permissionid = entity.PermissionID;
                    //}
                    //Check for entities if there are any with same permission id and future date

                    TableQuery<ExpirationLinksTableEntity> query1 = new TableQuery<ExpirationLinksTableEntity>()
                    .Where(TableQuery.CombineFilters(
                    TableQuery.GenerateFilterCondition("PermissionID", QueryComparisons.Equal, entity.PermissionID),
                    TableOperators.And,
                    TableQuery.GenerateFilterCondition("ExpirationDate", QueryComparisons.GreaterThan, DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK"))));

                    //TableQuery.GenerateFilterCondition("ExpirationDate", QueryComparisons.GreaterThan, DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK"))));
                    var entities1 = cloudTable.ExecuteQuery<ExpirationLinksTableEntity>(query1).ToList();
                    if (entities1.Count() == 0)
                    {
                        log.LogInformation($"Day: {entity.PartitionKey}, ID:{entity.RowKey}\tName:{entity.SharedByUser}\tDescription{entity.SharedWithUser}\tWebURL:{entity.WebURL}");
                        // CSOM code //  RemoveSPUserPermission(entity.SharedWithUser, entity.ItemURL, entity.WebURL, log);
                        await RemoveUserPermission(entity, permissionid, log);
                    }
                    // await _client.PostAsync("https://swxexpirationlinkentry.azurewebsites.net/api/ExpirationLinkEntry?", new StringContent("Hello"));
                    UpdateTable(entity, cloudTable, log);
                }
                log.LogInformation($"Connected to the table:");
                /////////////////////////////////////////////////////////////
            }
            catch (Exception x)
            {
                log.LogInformation($"Error: {x.Message}");
            }
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

            //public ExpirationLinksTableEntity(string sharedByUser, string sharedWithUser, string permissionLevel, DateTime expirationDate)
            //{
            //    SharedByUser = sharedByUser;
            //    SharedWithUser = sharedWithUser;
            //    PermissionLevel = permissionLevel;
            //    ExpirationDate = expirationDate;

            //}
            //public ExpirationLinksTableEntity() { }
        }

        public static void RemoveSPUserPermission(string sharedWithUser, string itemUrl, string webURL, ILogger log)
        {
            log.LogInformation("In permission remover");
            ContextProvider contextProvider = new ContextProvider(log);
            string listTitle = "Documents";
            Web currentWeb;
            using (var ctx = contextProvider.GetAppOnlyClientContext(webURL))
            {
                currentWeb = ctx.Web;
                ctx.Load(currentWeb);
                ctx.Load(ctx.Web, a => a.Lists);
                ctx.ExecuteQuery();
                Microsoft.SharePoint.Client.List list = ctx.Web.Lists.GetByTitle(listTitle);
                ctx.ExecuteQuery();
                log.LogInformation("Lib found");
                //string document = "log.txt";
                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = @"<View><Query><Where><Eq><FieldRef Name='FileRef' /><Value Type='Url'>" + itemUrl + @"</Value></Eq></Where> 
               </Query> 
                <ViewFields><FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' /></ViewFields> 
         </View>"
                };
                ctx.Load(list);
                log.LogInformation("List loaded");
                var items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                log.LogInformation("Items loaded");
                ctx.ExecuteQuery();
                log.LogInformation($"Items: {items.Count}");
                foreach (var item in items)
                {
                    //var user_group = web.SiteGroups.GetByName("Site Members");
                    log.LogInformation($"{sharedWithUser} ");
                    var user_group = currentWeb.SiteUsers.GetByLoginName("i:0#.f|membership|" + sharedWithUser.Trim(';'));
                    ctx.Load(user_group);
                    ctx.Load(items);
                    ctx.ExecuteQuery();
                    log.LogInformation($"Removing permissions {sharedWithUser.Trim(';')} ");
                    ctx.Load(item.RoleAssignments);
                    ctx.ExecuteQuery();

                    foreach (var assignments in item.RoleAssignments)
                    {
                        ctx.Load(assignments.Member);
                        ctx.ExecuteQuery();
                        if (assignments.Member.LoginName == user_group.LoginName)
                        {
                            item.RoleAssignments.GetByPrincipal(user_group).DeleteObject();
                            ctx.ExecuteQuery();
                        }
                    }
                    log.LogInformation($"Permissions removed: {user_group.Email} {DateTime.Now}");

                }
                log.LogInformation("\t Total Document = " + items.Count);
                string[] scopes = new string[] { $"https://graph.microsoft.com/Sites.FullControl.All" };
                var credential = new Azure.Identity.ClientSecretCredential(Environment.GetEnvironmentVariable("AADTenantId"), Environment.GetEnvironmentVariable("AADClientId"), Environment.GetEnvironmentVariable("AADClientSecret")
                );
                // This example works with Microsoft.Graph 4+
                var httpClientnew = Microsoft.Graph.GraphClientFactory.Create(new Microsoft.Graph.TokenCredentialAuthProvider(credential, scopes));
                // Create a new instance of GraphServiceClient with the authentication provider.

                Microsoft.Graph.GraphServiceClient graphClient = new Microsoft.Graph.GraphServiceClient(httpClientnew);
                log.LogInformation($"Created graph client");
            }
            string responseMessage = $"Hello from {currentWeb.Title}";
            log.LogInformation(responseMessage);
        }

        public static async Task<IDriveItemInviteCollectionPage> RemoveUserPermission(ExpirationLinksTableEntity entity, String PermissionID, ILogger log)
        {
            Microsoft.Graph.IDriveItemInviteCollectionPage result = null;
            try
            {
               // string[] scopes = new string[] { $"https://graph.microsoft.com/.default" };
                 string[] scopes = new string[] { Environment.GetEnvironmentVariable("GraphScope") };
                var credential = new Azure.Identity.ClientSecretCredential(Environment.GetEnvironmentVariable("AADTenantId"), Environment.GetEnvironmentVariable("AADClientId"), Environment.GetEnvironmentVariable("AADClientSecret"));
               // var credential = new Azure.Identity.ClientSecretCredential("973b3f01-ab0a-4620-b906-eb32095e50cc", "dd712a30-caac-45eb-b048-1cf1853b8dd8", "vSbPh-wRd-G4F865u58cD.Aj.ERqaUz-tP");
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

                await graphClient.Sites[entity.SiteID].Lists[entity.ListID].Items[entity.ItemID].DriveItem.Permissions[entity.PermissionID]
                    .Request()
                    .DeleteAsync();

                //result = await graphClient.Sites["76956606-5f0c-45e9-a4d0-7870674a3677"].Lists["7f036708-3d9a-46c9-a260-3b724b7e32d9"].Items["1"].DriveItem
                //            .Invite(recipients, requireSignIn, roles, sendInvitation, message, null, null)
                //            .Request()
                //            .PostAsync();
            }
            catch (Exception ex)
            {
                log.LogInformation($"{ex.Message}");
            }

            return result;
        }

        public static void UpdateTable(ExpirationLinksTableEntity entity,
           [Table("tblExpirationLinks", Connection = "AzureWebJobsStorage")] CloudTable cloudTable,
           ILogger log)
        {
            log.LogInformation($"Updating Azure table");
            try
            {
                entity.Expired = Convert.ToBoolean(true);
                entity.ExpirationDate = entity.ExpirationDate;
                entity.ETag = "*";
                TableOperation insertOrMergeOperation = TableOperation.Merge(entity);
                cloudTable.Execute(insertOrMergeOperation);

            }
            catch (Exception ex)
            {
                log.LogInformation($"Error: {ex.Message}");
            }

        }
        public static void UpdatePermissionIDInTable(ExpirationLinksTableEntity entity,
           [Table("tblExpirationLinks", Connection = "AzureWebJobsStorage")] CloudTable cloudTable,string permissionID,
           ILogger log)
        {
            log.LogInformation($"Updating Azure table");
            try
            {
                entity.PermissionID = permissionID;
                entity.ExpirationDate = entity.ExpirationDate;
                entity.ETag = "*";
                TableOperation insertOrMergeOperation = TableOperation.Merge(entity);
                cloudTable.Execute(insertOrMergeOperation);

            }
            catch (Exception ex)
            {
                log.LogInformation($"Error: {ex.Message}");
            }

        }
        public static async Task<String> GetUnresolvedUserPermId(ExpirationLinksTableEntity entity, ILogger log)
        {
            Microsoft.Graph.IDriveItemPermissionsCollectionPage result = null;
            string permissionId = "";
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


                result = await graphClient.Sites[entity.SiteID].Lists[entity.ListID].Items[entity.ItemID].DriveItem.Permissions
                    .Request().GetAsync();

                foreach (var test in result)
                {
                    if (test.GrantedToIdentities.Count() > 0)
                    {
                        var person = test.GrantedToIdentities.ToArray();
                        if (person[0].User.AdditionalData.ToArray()[0].Key == "email")
                        {
                            if (person[0].User.AdditionalData.ToArray()[0].Value.ToString() == entity.SharedWithUser)
                            {
                                permissionId = test.Id;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                log.LogInformation($"{ex.Message}");
            }
            return permissionId;
        }
    }
}
