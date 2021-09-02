using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Text;

namespace ConsoleApp2
{
    internal static class Program
    {
        private static async Task Main(string[] args)
        {
            // my unique organization's Azure DevOps uri
            const string uri = "https://dev.azure.com" + "/SWBC-FigWebDev";
            const string personalAccessToken = "k6negdnyagot2e35rlqlc4ovxeksmmmgv2ojn4olncq7eujbmu4a";


            var credentials = new VssBasicCredential(string.Empty, personalAccessToken);

            // connect to Azure DevOps
            var connection = new VssConnection(new Uri(uri), credentials);
            var workItemTrackingHttpClient = connection.GetClient<WorkItemTrackingHttpClient>();

            // get information about workitem
            const int workItemId = 209565;
            var workItem = workItemTrackingHttpClient.GetWorkItemAsync(workItemId).Result;

            var response = workItem.Url;
            Console.WriteLine(response);           
        }
    }
}
