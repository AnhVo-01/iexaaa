using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace IEXAAA.Dto
{
    public class GithubRelease
    {
        [JsonPropertyName("tag_name")]
        public string TagName { get; set; }

        [JsonPropertyName("html_url")]
        public string HtmlUrl { get; set; }

        public bool Prerelease { get; set; }

        public string Name { get; set; }

        public List<GithubAsset> Assets { get; set; }
    }
}
