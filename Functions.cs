﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.XPath;
using ExcelDna.Integration;
using ExcelDna.Registration;

namespace ImportFunctions
{
    public static class Functions
    {
        //// We will be using the single HttpClient from multiple threads,
        //// which is OK as long as we're not changing the default request headers.
        //static readonly HttpClient _httpClient;

        static Functions()
        {
        //    _httpClient = new HttpClient();
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
                IConfiguration config = Configuration.Default.WithDefaultLoader();
                IBrowsingContext context = BrowsingContext.New(config);
                IDocument document = await context.OpenAsync(url);

                var nodes = document.Body.SelectNodes(xpathQuery);

                if (nodes == null || nodes.Count == 0)
                    return "Error: No data found for the given XPath query";

                // return an object[] array with a single column containing the InnterText of the nodes
                var resultArray = new object[nodes.Count, 1];
                for (int i = 0; i < nodes.Count; i++)
                {
                    resultArray[i, 0] = nodes[i].TextContent;
                }
                return resultArray;
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
            [ExcelArgument(Description = "One-based index of the table or list to import from the HTML page. For example, 1 for the first table/list, 2 for the second, and so on.")]
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
                IConfiguration config = Configuration.Default.WithDefaultLoader();
                IBrowsingContext context = BrowsingContext.New(config);
                IDocument document = await context.OpenAsync(url);

                object result;
                if (dataType == "table")
                    result = ExtractTable(document, index);
                else
                    result = ExtractList(document, index);

                return result;
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

        //[ExcelFunction(Description = "Imports data from a given URL")]
        //public static async Task<object> HttpGet(string url)
        //{
        //    if (string.IsNullOrWhiteSpace(url))
        //    {
        //        return "Error: URL is required";
        //        // return ExcelError.ExcelErrorValue;
        //    }

        //    try
        //    {
        //        var response = await _httpClient.GetStringAsync(url);
        //        return response;
        //    }
        //    catch (HttpRequestException rex)
        //    {
        //        return $"Error: Unable to fetch data from the URL - {rex.Message}";
        //    }
        //    catch (Exception ex)
        //    {
        //        return $"Error: {ex.Message}";
        //    }
        //}

        static object ExtractTable(IDocument document, int indexOneBased)
        {
            var tables = document.Body.SelectNodes("//table");
            if (tables == null || tables.Count < indexOneBased)
                return "Error: Table not found";

            var table = (IElement)tables[indexOneBased - 1];

            var results = new List<List<string>>();
            foreach (var row in table.SelectNodes(".//tr").Cast<IElement>())
            {
                var rowResult = new List<string>();
                foreach (var cell in row.SelectNodes(".//th|.//td").Cast<IElement>())
                {
                    rowResult.Add(cell.TextContent);
                }
                results.Add(rowResult);
            }

            if (results.Count == 0 || results[0].Count == 0)
                return "Error: No data found in the table";

            // Convert results to a 2D object array
            var resultArray = new object[results.Count, results[0].Count];
            for (int i = 0; i < results.Count; i++)
            {
                for (int j = 0; j < results[i].Count; j++)
                {
                    resultArray[i, j] = results[i][j];
                }
            }
            return resultArray;
        }

        static object ExtractList(IDocument document, int indexOneBased)
        {
            var lists = document.Body.SelectNodes("//ul | //ol");
            if (lists == null || lists.Count < indexOneBased)
                return "Error: List not found";

            var list = (IElement)lists[indexOneBased - 1];

            var results = new List<string>();
            foreach (var item in list.SelectNodes(".//li"))
            {
                results.Add(item.TextContent);
            }

            // Convert results to a 2D object array with a single column
            var resultArray = new object[results.Count, 1];
            for (int i = 0; i < results.Count; i++)
            {
                resultArray[i, 0] = results[i];
            }

            return resultArray;
        }
    }
}
