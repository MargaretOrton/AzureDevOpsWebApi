using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CopyWorkItems
{
    class Program
    {
        private static WorkItemTrackingHttpClient WitClient;

        private static readonly string AzureDevOpsOrgUrl = "https://dev.azure.com/<enter your org name>/";

        // Create an access token as described here:
        // https://docs.microsoft.com/en-us/azure/devops/organizations/accounts/use-personal-access-tokens-to-authenticate?view=azure-devops&tabs=preview-page
        private static readonly  string AzureDevOpsUserToken = "ACCESS TOKEN";        
        private static readonly string MasterTeamProjectName = "PROJECT NAME TO COPY ITEMS FROM"; 
        private static readonly string NewTeamProjectName = "PROJECT NAME TO COPY ITEMS TO";
        private static readonly string QueryPath = "Shared Queries/sharedquery"; //change to your query path

        private static Dictionary<int, int> _oldToNewWorkItemIdMap;
        private static Dictionary<int, string> _newItemIdToUrlMap;

        static void Main(string[] args)
        {
            _oldToNewWorkItemIdMap = new Dictionary<int /*original work item id*/, int /*new work item id*/>();
            _newItemIdToUrlMap = new Dictionary<int /*new work item id*/, string /*new work item url*/>();

            try
            {
                var connection = new VssConnection(new Uri(AzureDevOpsOrgUrl), new VssBasicCredential(string.Empty, AzureDevOpsUserToken));
                WitClient = connection.GetClient<WorkItemTrackingHttpClient>();
                              

                List<int> workItemIds = GetStoredQueryWorkItemIds(MasterTeamProjectName, QueryPath);
                
                var workItems = WitClient.GetWorkItemsAsync(workItemIds, expand: WorkItemExpand.Relations).Result;
                
                // copy all items into new project without linking
                foreach (var workItem in workItems)
                {
                    Dictionary<string, object> fields = GetFieldsCopy(workItem);

                    var newItem = AddWorkItem(NewTeamProjectName, workItem.Fields["System.WorkItemType"].ToString(), fields);
                   
                    _oldToNewWorkItemIdMap.Add(workItem.Id.Value, newItem.Id.Value);
                    _newItemIdToUrlMap.Add(newItem.Id.Value, newItem.Url);
                }

                // set parent-child links to newly created items
                foreach (var workItem in workItems)
                {
                    if (workItem.Relations == null || !workItem.Relations.Any()) continue;

                    

                    foreach (var relatedItem in workItem.Relations)
                    {
                        var childOldIdStr = relatedItem.Url.Split('/').Last();

                        if (!int.TryParse(childOldIdStr, out int childOldId)) continue;
                        if (!_oldToNewWorkItemIdMap.ContainsKey(childOldId)) continue;

                        var newItemId = _oldToNewWorkItemIdMap[childOldId];
                        if (!_newItemIdToUrlMap.ContainsKey(newItemId)) continue;

                        var newItemUrl = _newItemIdToUrlMap[newItemId];

                        JsonPatchDocument patchDocument = new JsonPatchDocument();
                        patchDocument.Add(new JsonPatchOperation()
                        {
                            Operation = Operation.Add,
                            Path = "/relations/-",
                            Value = new
                            {
                                rel = relatedItem.Rel,
                                url = newItemUrl,
                                attributes = new
                                {
                                    name = relatedItem.Attributes["name"],
                                    isLocked = relatedItem.Attributes["isLocked"] // you must be an administrator to lock a link
                                }
                            }
                        });

                        var itemToUpdateId = _oldToNewWorkItemIdMap[workItem.Id.Value];
                        UpdateWorkItemLink(itemToUpdateId, patchDocument);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                if (ex.InnerException != null) Console.WriteLine("Detailed Info: " + ex.InnerException.Message);
                Console.WriteLine("Stack:\n" + ex.StackTrace);
            }
        }

        private static Dictionary<string, object> GetFieldsCopy(WorkItem workItem)
        {
            Dictionary<string, object> fields = new Dictionary<string, object>();

            fields.Add("System.TeamProject", NewTeamProjectName);
            fields.Add("System.AreaPath", NewTeamProjectName);
            fields.Add("System.IterationPath", NewTeamProjectName);
            fields.Add("System.WorkItemType", workItem.Fields["System.WorkItemType"]);
            fields.Add("System.Title", workItem.Fields["System.Title"]);

            if (workItem.Fields.ContainsKey("System.Description"))
                fields.Add("System.Description", workItem.Fields["System.Description"]);
            
            return fields;
        }

        static List<int> GetStoredQueryWorkItemIds(string project, string queryPath)
        {
            var query = WitClient.GetQueryAsync(project, queryPath, QueryExpand.Wiql).Result;

            string wiqlStr = query.Wiql;

            return GetQueryResult(wiqlStr, project);
        }

        static List<int> GetQueryResult(string wiqlStr, string teamProject)
        {
            WorkItemQueryResult result = RunQueryByWiql(wiqlStr, teamProject);

            if (result != null)
            {
                if (result.WorkItems != null) // this is Flat List 
                    return (from wis in result.WorkItems select wis.Id).ToList();
                else Console.WriteLine("There is no query result");
            }

            return new List<int>();
        }

        static WorkItemQueryResult RunQueryByWiql(string wiqlStr, string teamProject)
        {
            Wiql wiql = new Wiql();
            wiql.Query = wiqlStr;

            if (teamProject == "")
                return WitClient.QueryByWiqlAsync(wiql).Result;
            else 
                return WitClient.QueryByWiqlAsync(wiql, teamProject).Result;
        }

        static WorkItem AddWorkItem(string projectName, string workItemTypeName, Dictionary<string, object> fields)
        {
            JsonPatchDocument patchDocument = new JsonPatchDocument();

            foreach (var key in fields.Keys)
                patchDocument.Add(new JsonPatchOperation()
                {
                    Operation = Operation.Add,
                    Path = "/fields/" + key,
                    Value = fields[key]
                });

            return WitClient.CreateWorkItemAsync(patchDocument, projectName, workItemTypeName).Result;

        }

        static WorkItem UpdateWorkItemLink(int workItemId, JsonPatchDocument patchDocument)
        {
            return WitClient.UpdateWorkItemAsync(patchDocument, workItemId).Result;
        }
    }
}
