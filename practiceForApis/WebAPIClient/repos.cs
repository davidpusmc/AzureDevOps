using System;
using System.Text.Json.Serialization;

namespace WebAPIClient
{
    public class Repository
    {
        //gets json by given apis and sets them for use with program.cs
        [JsonPropertyName("name")]
        public string Name { get; set;}
        [JsonPropertyName("id")]
        public int Id { get; set;}
        [JsonPropertyName("description")]
        public string Description {get; set;}
        [JsonPropertyName("html_url")]
        public Uri GitHubHomeUrl {get; set;}
        [JsonPropertyName("homepage")]
        public Uri Homepage { get; set; }
        [JsonPropertyName("watchers")]
        public int Watchers { get; set; }

        [JsonPropertyName("pushed_at")]
        public DateTime LastPushUtc { get; set; }

        public DateTime LastPush => LastPushUtc.ToLocalTime();
    }
}