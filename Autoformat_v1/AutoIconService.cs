using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace Autoformat_v1
{
    public static class AutoIconService
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private static JArray _iconsDatabase = null;
        private const string ICONS_JSON_URL = "https://yash-jadhav-gd.github.io/Icon_finder/assets/icons.json";
        private const string ICONS_BASE_URL = "https://yash-jadhav-gd.github.io/Icon_finder/assets/icons/";

        private static readonly HashSet<string> StopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "a", "an", "the", "and", "or", "but", "is", "are", "was", "were",
            "in", "on", "at", "to", "for", "with", "about", "against", "between",
            "into", "through", "during", "before", "after", "above", "below",
            "from", "up", "down", "of", "off", "over", "under", "again", "further",
            "then", "once", "here", "there", "when", "where", "why", "how", "all",
            "any", "both", "each", "few", "more", "most", "other", "some", "such",
            "no", "nor", "not", "only", "own", "same", "so", "than", "too", "very"
        };

        private static async Task EnsureDatabaseLoadedAsync()
        {
            if (_iconsDatabase != null) return;
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12 | System.Net.SecurityProtocolType.Tls11 | System.Net.SecurityProtocolType.Tls;
                var response = await _httpClient.GetStringAsync(ICONS_JSON_URL);
                _iconsDatabase = JArray.Parse(response);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error loading icons.json: {ex.Message}");
            }
        }

        public static string GetBestIconSvg(string text)
        {
            // Run synchronously for COM Interop compatibility in VSTO
            return Task.Run(() => GetBestIconSvgAsync(text)).GetAwaiter().GetResult();
        }

        public static async Task<string> GetBestIconSvgAsync(string text)
        {
            await EnsureDatabaseLoadedAsync();
            if (_iconsDatabase == null || _iconsDatabase.Count == 0) return null;

            // Clean text and extract keywords
            var words = Regex.Matches(text, @"\b[a-zA-Z]{3,}\b")
                             .Cast<Match>()
                             .Select(m => m.Value.ToLowerInvariant())
                             .Where(w => !StopWords.Contains(w))
                             .ToList();

            if (words.Count == 0) words.Add("circle"); // Fallback if no keywords

            JToken bestMatch = null;
            int highestScore = -1;

            foreach (var icon in _iconsDatabase)
            {
                int score = 0;
                string label = icon["label"]?.ToString().ToLowerInvariant() ?? "";
                var tags = icon["tags"]?.Select(t => t.ToString().ToLowerInvariant()).ToList() ?? new List<string>();

                foreach (var word in words)
                {
                    if (label.Contains(word)) score += 5; // Label matches are strong
                    if (tags.Contains(word)) score += 2; // Exact tag matches
                    else if (tags.Any(t => t.Contains(word))) score += 1; // Partial tag matches
                }

                if (score > highestScore)
                {
                    highestScore = score;
                    bestMatch = icon;
                }
            }

            if (bestMatch == null) bestMatch = _iconsDatabase.First(); // Fallback to first if nothing matches

            string fileName = bestMatch["file"]?.ToString();
            if (string.IsNullOrEmpty(fileName)) return null;

            try
            {
                string svgUrl = ICONS_BASE_URL + fileName;
                string svgContent = await _httpClient.GetStringAsync(svgUrl);
                return svgContent;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error fetching SVG: {ex.Message}");
                return null;
            }
        }

        public static string SaveSvgToTempFile(string svgContent)
        {
            if (string.IsNullOrEmpty(svgContent)) return null;
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".svg");
            File.WriteAllText(tempPath, svgContent);
            return tempPath;
        }
    }
}
