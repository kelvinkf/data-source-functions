using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using Microsoft.Graph.Auth;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Extensions.Options;

namespace Plumsail.DataSource.SharePoint
{
    public class ListData
    {
        private readonly Settings.ListData _settings;
        private readonly GraphServiceClientProvider _graphProvider;

        public ListData(IOptions<Settings.AppSettings> settings, GraphServiceClientProvider graphProvider)
        {
            _settings = settings.Value.ListData;
            _graphProvider = graphProvider;
        }

        [FunctionName("SharePoint-ListData")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("ListData is requested.");

            var graph = _graphProvider.Create();
            var list = await graph.GetListAsync(_settings.SiteUrl, _settings.ListName);
            string token = req.Query["n"];

            var queryOptions = new List<QueryOption>()
            {
                new QueryOption("select", "id"),
                new QueryOption("expand", "fields(select=Title,Email,Description,Reimbursement_x0020_Type,Reimbursement_x0020_Amount,Token)"),
                new QueryOption("filter", $"startswith(fields/Token, '{token}')")
                //new QueryOption("filter", "fields/Token/any(a:a+eq+'63d8c29f-3c41-47a3-a282-3d66fe8668b6')")
                //new QueryOption("filter", "fields/Token eq '63d8c29f-3c41-47a3-a282-3d66fe8668b6')")
            };
            var itemsPage = await list.Items
                .Request(queryOptions)
                .GetAsync();
            var items = new List<ListItem>(itemsPage);

            while (itemsPage.NextPageRequest != null)
            {
                itemsPage = await itemsPage.NextPageRequest.GetAsync();
                items.AddRange(itemsPage);
            }

            return new OkObjectResult(items);
        }
    }
}
