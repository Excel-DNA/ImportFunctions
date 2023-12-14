using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using ExcelDna.Integration;
using ExcelDna.Registration;
using HtmlAgilityPack;

namespace ImportFunctions
{
    public static class Functions
    {
        // We will be using the single HttpClient from multiple threads,
        // which is OK as long as we're not changing the default request headers.
        static readonly HttpClient _httpClient;

        static Functions()
        {
            _httpClient = new HttpClient();
            ServicePointManager.SecurityProtocol =
                      SecurityProtocolType.Tls |
                      SecurityProtocolType.Tls11 |
                      SecurityProtocolType.Tls12 |
                      SecurityProtocolType.Tls13;
        }

        [ExcelAsyncFunction(Description = "Imports data from a given URL using an XPath query")]
        public static async Task<object> ImportXml(string url, string xpathQuery)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return "Error: URL is required";
                // return ExcelError.ExcelErrorValue;
            }

            if (string.IsNullOrWhiteSpace(xpathQuery))
            {
                return "Error: XPath query is required";
                // return ExcelError.ExcelErrorValue;
            }

            try
            {
                var response = await _httpClient.GetStringAsync(url);
                var doc = new HtmlDocument();
                doc.LoadHtml(response);

                var node = doc.DocumentNode.SelectSingleNode(xpathQuery);
                return node?.InnerText ?? "Error: No data found for the given XPath query";
            }
            catch (HttpRequestException rex)
            {
                return $"Error: Unable to fetch data from the URL - {rex.Message}";
            }
            catch (XmlException xex)
            {
                return $"Error: Invalid XML data - {xex.Message}";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Imports data from a table or list within an HTML page")]
        public static async Task<object> ImportHtml(
            [ExcelArgument(Description = "URL of the HTML page to scrape data from. The URL must start with either http or https.")]
            string url,
            [ExcelArgument(Description = "Type of data to import. Accepts either 'table' for HTML tables or 'list' for HTML lists (ul/ol).")]
            string dataType,
            [ExcelArgument(Description = "Zero-based index of the table or list to import from the HTML page. For example, 0 for the first table/list, 1 for the second, and so on.")]
            int index)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return "Error: URL is required";
                // return ExcelError.ExcelErrorValue;
            }

            if (dataType != "table" && dataType != "list")
            {
                return "Error: Data type must be 'table' or 'list'";
                // return ExcelError.ExcelErrorValue;
            }

            try
            {
                var response = await _httpClient.GetStringAsync(url);
                var doc = new HtmlDocument();
                doc.LoadHtml(response);

                if (dataType == "table")
                    return ExtractTable(doc, index);
                else
                    return ExtractList(doc, index);
            }
            catch (HttpRequestException rex)
            {
                return $"Error: Unable to fetch data from the URL - {rex.Message}";
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
        }

        static object ExtractTable(HtmlDocument doc, int index)
        {
            var tables = doc.DocumentNode.SelectNodes("//table");
            if (tables == null || tables.Count <= index)
                return "Error: Table not found";

            var table = tables[index];
            var sb = new StringBuilder();

            foreach (var row in table.SelectNodes("tr"))
            {
                foreach (var cell in row.SelectNodes("th|td"))
                {
                    sb.Append(cell.InnerText.Trim());
                    sb.Append("\t"); // Tab-separated values
                }
                sb.AppendLine(); // New line at the end of each row
            }

            return sb.ToString();
        }

        static object ExtractList(HtmlDocument doc, int index)
        {
            var lists = doc.DocumentNode.SelectNodes("//ul | //ol");
            if (lists == null || lists.Count <= index)
                return "Error: List not found";

            var list = lists[index];
            var sb = new StringBuilder();

            foreach (var item in list.SelectNodes("li"))
            {
                sb.AppendLine(item.InnerText.Trim());
            }

            return sb.ToString();
        }
    }
}
