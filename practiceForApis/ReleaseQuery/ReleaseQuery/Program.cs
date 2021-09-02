using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
 

namespace ReleaseQuery
{
    internal class Program
    {
        private static List<string> listResult;
        

        private static void Main(string[] args)
        {
            string CollectionUri = "https://dev.azure.com/SWBC-FigWebDev";
            string adoPat = "k6negdnyagot2e35rlqlc4ovxeksmmmgv2ojn4olncq7eujbmu4a";
            List<WorkItem> workItems;
            DateTime startDateTime = DateTime.Now;
            String teamProjectName = "DevOps";//changed project name item runs but query fails updated 5/28/2021

            VssConnection vssConnection = new VssConnection(new Uri(CollectionUri),
                                                            new VssBasicCredential(string.Empty, adoPat));

            WorkItemTrackingHttpClient witClient = vssConnection.GetClient<WorkItemTrackingHttpClient>();

            WorkItemQueryResult wiqResult = witClient.QueryByWiqlAsync(
                new Wiql
                {
                    Query = "SELECT * "
                    + "FROM WorkItems "
                    + "WHERE [System.TeamProject] = '" + teamProjectName + "' "
                    + "AND ( "
                    + "  [System.WorkItemType] = 'Product Backlog Item' "
                    + "  OR [System.WorkItemType] = 'Bug' "
                    + ") "
                    + "AND [System.State] = 'Done' "
                    + "AND [Microsoft.VSTS.Common.ImpDate] >= '" + startDateTime.AddDays(-7).ToString("MM/dd/yyyy") + "' "
                    + "AND [Microsoft.VSTS.Common.ImpDate] <= '" + startDateTime.ToString("MM/dd/yyyy") + "' "
                    + "AND ( "
                    + "  [SWBC.ApplicationType] = 'Auto' "
                    + "  OR [SWBC.ApplicationType] = 'GAP' "
                    + "  OR [SWBC.ApplicationType] = 'Mortgage' "
                    + "  OR [SWBC.ApplicationType] = 'Accounting' "
                    + "  OR [SWBC.ApplicationType] = 'Stone Eagle' "
                    + ") "
                    + "AND [System.ExternalLinkCount] > 0 "
                    + "AND ( "
                    + "  [SWBC.AS400Task] <> 'NONE' "
                    + "  AND [SWBC.AS400Task] <> 'NA' "
                    + "  AND [SWBC.AS400Task] <> 'N/A' "
                    + ") "
                },false, null, false
            ).Result;

            if (wiqResult.WorkItems.Any())
            {
                workItems = witClient.GetWorkItemsAsync(wiqResult.WorkItems.Select(wir => wir.Id)).Result;
                
            }
            else
            {
                workItems = new List<WorkItem>();
                Console.WriteLine(workItems.ToString());
                
                
            }
            ;

            // ---
            listResult = new List<String>();
            string resultsAdoQuery = String.Join("\t", new string[] { "ID", "Title", "AS400 Task", "SWBC Application Type", "Implementation Date", "Work Item Type", "State", }) + Environment.NewLine;

            foreach (WorkItem workItem in workItems)
            {
                listResult.Add(workItem.Fields.GetValueOrDefault("SWBC.AS400Task").ToString().Trim());

                resultsAdoQuery += String.Join("\t", new string[]
                {
                    workItem.Id.ToString().Trim(),

                    workItem.Fields.GetValueOrDefault("System.Title", "<missing>").ToString().Trim()
                }
                
                );
                

            }
        }
    }
}