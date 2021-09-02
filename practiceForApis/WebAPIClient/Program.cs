using System;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Collections.Generic;
using System.Text.Json;

namespace WebAPIClient
{
    class Program
    {
        
        private static readonly HttpClient client = new HttpClient();//will handle requests and responses
        static async Task Main(string[] args)//captures results and writes each repo property
        {
            var repositories = await ProcessRepositories();
            foreach (var repo in repositories)
            {
                Console.WriteLine(repo.Name);
                Console.WriteLine(repo.Description);
                Console.WriteLine(repo.GitHubHomeUrl);
                Console.WriteLine(repo.Homepage);
                Console.WriteLine(repo.Watchers);
                Console.WriteLine(repo.LastPush + Environment.NewLine);
            }
        }
        private static async Task<List<Repository>> ProcessRepositories()
        {
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));//which content types are accepted.
            client.DefaultRequestHeaders.Add("User-Agent", ".NET Foundation Repository Reporter");//headers necessary to retrieve info from source

            var streamTask = client.GetStreamAsync("https://api.github.com/orgs/dotnet/repos");//starts the task for webrequest and returns with response
            var  repositories = await JsonSerializer.DeserializeAsync<List<Repository>>(await streamTask);
            return repositories;

            
        }
    }
}
