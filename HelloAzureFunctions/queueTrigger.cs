
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Table;
using PnP.Framework.Provisioning.Model;
using PnP.Framework.Provisioning.ObjectHandlers;
using PnP.Framework.Provisioning.Providers.Xml;
using System.Reflection;

namespace HelloAzureFunctions;

public class queueTrigger
{
    private readonly ILogger<queueTrigger> _logger;

    public queueTrigger(ILogger<queueTrigger> logger)
    {
        _logger = logger;
    }

    [Function(nameof(queueTrigger))]
    public async Task RunAsync([QueueTrigger("myqueue-items", Connection = "AzureWebJobsStorage")] Azure.Storage.Queues.Models.QueueMessage message)
    {
        try
        {

            _logger.LogWarning("Function triggered...");
            string accessToken = null;
            try
            {
                _logger.LogWarning("Attempting to retrieve access token...");
                accessToken = await CertificateLoader.GetAccessTokenAsync();

                if (string.IsNullOrEmpty(accessToken))
                {
                    _logger.LogError("Access token is null or empty...!");
                    throw new InvalidOperationException("Access token retrieval failed...!");
                }
                _logger.LogWarning("Access token successfully retrived...");
            }
            catch (Exception ex)
            {
                _logger.LogError("Error while accessing token : ", ex.Message);
                throw;
            }
            string SiteURL = Environment.GetEnvironmentVariable("SiteURL");
            //_logger.LogInformation($"SiteURL: {SiteURL}");
            if (string.IsNullOrEmpty(SiteURL))
            {
                _logger.LogError("SiteUrl is Missing...!");
                throw new InvalidOperationException("Missing SiteUrl...!");
            }

            var storageAccount = CloudStorageAccount.Parse(Environment.GetEnvironmentVariable("AzureWebJobsStorage"));
            var tableClient = storageAccount.CreateCloudTableClient();
            var table = tableClient.GetTableReference("ChangeTokenCache");
            await table.CreateIfNotExistsAsync();

            string partitionKey = "ChangeToken";
            string rowKey = "LastProcessed";

            // getting last change token
            var retrieveOperation = TableOperation.Retrieve<ChangeTokenEntity>(partitionKey, rowKey);
            var retrievedResult = await table.ExecuteAsync(retrieveOperation);
            var tokenEntity = retrievedResult.Result as ChangeTokenEntity;

            if (tokenEntity == null)
            {
                _logger.LogWarning("No previous change token found. Starting fresh.");
            }
            else
            {
                _logger.LogWarning($"Token Entity : {tokenEntity}");
            }
            using (var context = new ClientContext(SiteURL))
            {
                context.ExecutingWebRequest += (sender, e) => {
                    e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                };

                var list = context.Web.Lists.GetByTitle("Client List");
                ChangeQuery cq = new ChangeQuery(false, false)
                {
                    Item = true,
                    Add = true,
                    Update = false,
                    DeleteObject = false
                };

                if (tokenEntity != null && !string.IsNullOrEmpty(tokenEntity.ChangeToken))
                {
                    cq.ChangeTokenStart = new ChangeToken { StringValue = tokenEntity.ChangeToken };
                    _logger.LogWarning("Using stored change token: {Token}", tokenEntity.ChangeToken);
                }

                var changes = list.GetChanges(cq);
                context.Load(changes);
                try
                {
                    context.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    _logger.LogError("Error executing SharePoint query...!", ex.Message);
                    throw;
                }
                ChangeToken lastChangeToken = null;
                foreach (Change change in changes)
                {
                    if (change is ChangeItem changeItem)
                    {
                        switch (change.ChangeType)
                        {
                            case Microsoft.SharePoint.Client.ChangeType.Add:
                                _logger.LogWarning("Item was Added...");
                                Microsoft.SharePoint.Client.ListItem item = list.GetItemById(changeItem.ItemId);
                                context.Load(item);
                                try
                                {
                                    context.ExecuteQuery();
                                    string clientName = item["ClientName"]?.ToString() ?? "No Client Name";
                                    string clientNumber = item["ClientNumber"]?.ToString() ?? "No Client Number";
                                    string created = item["Created"]?.ToString() ?? "Unknown Date";
                                    //string assignedTo = item["AssignedTo"]?.ToString() ?? "Unassigned";

                                    _logger.LogWarning("Item Details - Client Name: {ClientName}, Client Number: {ClientNumber}, Created: {Created}", clientName, clientNumber, created);
                                    var groupBody = new Microsoft.Graph.Group
                                    {
                                        DisplayName = clientName,
                                        Description = "Site " + clientName,
                                        GroupTypes = new List<string> { "Unified" },
                                        MailEnabled = true,
                                        MailNickname = clientNumber,
                                        SecurityEnabled = false,
                                        Visibility = "Private"
                                    };

                                    var graphClient = CertificateLoader.GetGraphClient();
                                    var createdGroup = await graphClient.Groups.Request().AddAsync(groupBody);
                                    _logger.LogWarning($"Created group id: {createdGroup.Id}");
                                    var memberEmails = new List<string> { "arunjatak@dgneaseteq.onmicrosoft.com" };
                                    foreach (var email in memberEmails)
                                    {
                                        // Get the user object by email
                                        var user = await graphClient.Users[email].Request().GetAsync();

                                        // Add user as an owner
                                        await graphClient.Groups[createdGroup.Id].Owners.References
                                            .Request()
                                            .AddAsync(user);
                                    }
                                    _logger.LogWarning("Group Owners Added");

                                    foreach (var email in memberEmails)
                                    {
                                        // Get the user object by email
                                        var user = await graphClient.Users[email].Request().GetAsync();

                                        // Add user as an member
                                        await graphClient.Groups[createdGroup.Id].Members.References
                                             .Request()
                                             .AddAsync(user);
                                    }
                                    _logger.LogWarning("Group Members Added");
                                    await Task.Delay(TimeSpan.FromSeconds(15));
                                    var site = await graphClient.Groups[createdGroup.Id].Sites["root"].Request().GetAsync();
                                    var siteUrl = site.WebUrl;
                                    Console.WriteLine($"siteUrl: {siteUrl}");
                                    _logger.LogWarning($"siteUrl: {siteUrl}");
                                    using (var siteContext = new ClientContext(siteUrl))
                                    {
                                        siteContext.ExecutingWebRequest += (sender, e) =>
                                        {
                                            e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
                                        };
                                        _logger.LogWarning($"Configuring : {siteUrl}");
                                        var lists = siteContext.Web.Lists;
                                        siteContext.Load(lists,
                                            l => l.Include(
                                                x => x.Title,
                                                x => x.BaseTemplate,
                                                x => x.Hidden,
                                                x => x.RootFolder.ServerRelativeUrl
                                                )
                                            );
                                        siteContext.ExecuteQuery();
                                        var docLibrary = lists.FirstOrDefault(l => l.BaseTemplate == 101 && l.Title == "Documents");
                                        if (docLibrary != null)
                                        {
                                            _logger.LogWarning($"Documents Library : {docLibrary.Title}");
                                            var branches = GetBranch(SiteURL, accessToken, clientNumber);

                                            foreach (var branch in branches)
                                            {
                                                _logger.LogWarning($"Branch {branch.BranchName}");
                                                string branchDocLib;
                                                string branchPlannerPlan;
                                                // Create new document library
                                                ListCreationInformation lci = new ListCreationInformation
                                                {
                                                    Title = branch.BranchName,
                                                    Description = "DocLib",
                                                    TemplateType = 101
                                                };

                                                var newLib = siteContext.Web.Lists.Add(lci);
                                                siteContext.Load(newLib);
                                                try
                                                {
                                                    siteContext.ExecuteQuery();
                                                    _logger.LogWarning($"{branch.BranchName} Document Library Created Successfully...");
                                                }
                                                catch (Exception ex) { _logger.LogError($"Error While create document libraray : {ex.Message}"); }

                                                branchDocLib = siteUrl + "/" + branch.BranchName;
                                                await Task.Delay(TimeSpan.FromSeconds(5));

                                                var branchPlan = new PlannerPlan
                                                {
                                                    Owner = createdGroup.Id,
                                                    Title = branch.BranchName
                                                };
                                                var branchResult = await graphClient.Planner.Plans.Request().AddAsync(branchPlan);
                                                branchPlannerPlan = "https://planner.cloud.microsoft/webui/plan/" + branchResult.Id + "/view";
                                                _logger.LogWarning($"Planner plan created successfully for branch, Plan Name: {branchResult.Title}, Plan Id: {branchResult.Id}");
                                                siteContext.Load(newLib, l => l.ContentTypesEnabled, l => l.ContentTypes);
                                                try
                                                {
                                                    siteContext.ExecuteQuery();
                                                }
                                                catch (Exception ex)
                                                {
                                                    _logger.LogWarning($"Error while loading content type : ",ex.Message);
                                                }

                                                // Enable content types
                                                if (!newLib.ContentTypesEnabled)
                                                {

                                                    try
                                                    {
                                                        newLib.ContentTypesEnabled = true;
                                                        newLib.Update();
                                                        siteContext.ExecuteQuery();
                                                        _logger.LogWarning($"Content type is enabled in {branch.BranchName}");
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        _logger.LogWarning("Error while enabling content type...!", ex);
                                                    }
                                                    Console.WriteLine("Content types enabled on " + branch.BranchName + " library.");
                                                    await Task.Delay(TimeSpan.FromSeconds(3));

                                                    var defaultCT = newLib.ContentTypes.FirstOrDefault(ct => ct.Name == "Document");
                                                    if (defaultCT != null)
                                                    {
                                                        defaultCT.DeleteObject();
                                                        try
                                                        {
                                                            siteContext.ExecuteQuery();
                                                            _logger.LogWarning($"Default '{defaultCT.Name}' content type removed.");
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            _logger.LogWarning("Error While deleting default content type", ex);
                                                        }
                                                    }

                                                    string contentTypeId = "0x01010097E7D4CF65DFB04EA47382EADE69C850".Trim();
                                                    int retry = 0;
                                                    while (retry < 10)
                                                    {
                                                        try
                                                        {
                                                            var subscriber = new Microsoft.SharePoint.Client.Taxonomy.ContentTypeSync.ContentTypeSubscriber(siteContext);
                                                            siteContext.Load(subscriber);
                                                            siteContext.ExecuteQuery();
                                                            subscriber.SyncContentTypesFromHubSite2(siteUrl, new List<string> { contentTypeId });
                                                            await Task.Delay(5000);
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            _logger.LogWarning("Error While subscribe content type", ex.Message);
                                                        }

                                                        try
                                                        {
                                                            siteContext.Load(siteContext.Web.ContentTypes);
                                                            siteContext.ExecuteQuery();
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            _logger.LogWarning("Error While load content type : ", ex.Message);
                                                        }
                                                        var importedCT = siteContext.Web.ContentTypes.FirstOrDefault(ct => ct.Id.StringValue.StartsWith(contentTypeId));
                                                        if (importedCT != null)
                                                        {
                                                            newLib.ContentTypes.AddExistingContentType(importedCT);
                                                            newLib.Update();
                                                            try
                                                            {
                                                                siteContext.ExecuteQuery();
                                                                _logger.LogWarning($"Added content type '{importedCT.Name}' to the library.");
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                _logger.LogWarning("Error While adding content type : ",ex.Message);
                                                            }
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            _logger.LogWarning("Waiting for content type to sync...");
                                                            retry++;
                                                            await Task.Delay(5000);
                                                        }
                                                    }
                                                }

                                                // Set default field value
                                                try
                                                {
                                                    Microsoft.SharePoint.Client.Field clientField = newLib.Fields.GetByInternalNameOrTitle("ClientNumber1");
                                                    clientField.DefaultValue = clientNumber;
                                                    clientField.Update();
                                                    siteContext.Load(clientField);
                                                    siteContext.ExecuteQuery();
                                                    _logger.LogWarning("Default column set to the document library.");
                                                }
                                                catch (Exception ex)
                                                {
                                                    _logger.LogError("Error While setting default column value.", ex.Message);
                                                }


                                                // Create folders
                                                string[] folderNames = ["Folder1", "Folder2"];
                                                foreach (var folder in folderNames)
                                                {
                                                    var itemCreateInfo = new ListItemCreationInformation
                                                    {
                                                        UnderlyingObjectType = Microsoft.SharePoint.Client.FileSystemObjectType.Folder,
                                                        LeafName = folder
                                                    };
                                                    Microsoft.SharePoint.Client.ListItem folderItem = newLib.AddItem(itemCreateInfo);
                                                    folderItem["Title"] = folder;
                                                    folderItem.Update();
                                                    siteContext.ExecuteQuery();
                                                }
                                                _logger.LogWarning("Folder created in document library...");



                                                //Create site page
                                                try
                                                {
                                                    Microsoft.SharePoint.Client.List Library = siteContext.Web.Lists.GetByTitle("site pages");
                                                    siteContext.Load(siteContext.Web, w => w.ServerRelativeUrl);
                                                    siteContext.ExecuteQuery();
                                                    string serverRelativeUrl = $"{siteContext.Web.ServerRelativeUrl.TrimEnd('/')}/SitePages/" + branch.BranchName + ".aspx";
                                                    Microsoft.SharePoint.Client.ListItem oItem = Library.RootFolder.Files
                                                        .AddTemplateFile(serverRelativeUrl, TemplateFileType.ClientSidePage)
                                                        .ListItemAllFields;

                                                    oItem["ContentTypeId"] = "0x0101009D1CB255DA76424F860D91F20E6C4118";
                                                    oItem["Title"] = System.IO.Path.GetFileNameWithoutExtension(branch.BranchName + ".aspx");
                                                    oItem["ClientSideApplicationId"] = "b6917cb1-93a0-4b97-a84d-7cf49975d4ec";
                                                    oItem["PageLayoutType"] = "Article";
                                                    oItem["PromotedState"] = "0";
                                                    oItem["CanvasContent1"] = "<div></div>";
                                                    oItem["BannerImageUrl"] = "/_layouts/15/images/sitepagethumbnail.png";

                                                    oItem.Update();
                                                    siteContext.Load(oItem, item => item.Id);
                                                    siteContext.ExecuteQuery();

                                                    _logger.LogWarning("Successfully created modern page in library, Page ID: " + oItem.Id);
                                                }
                                                catch (Exception ex)
                                                {
                                                    _logger.LogError("Error Occured while creating page.", ex.Message);
                                                }
                                                await Task.Delay(TimeSpan.FromSeconds(5));

                                                //Add webpart to page
                                                string relativeUrl = "/sites/Arun/Shared Documents/TeamSiteStructure.xml";
                                                var file = context.Web.GetFileByServerRelativeUrl(relativeUrl);
                                                var fileStream = file.OpenBinaryStream();
                                                context.Load(file);
                                                context.ExecuteQuery();

                                                using (var memoryStream = new MemoryStream())
                                                {
                                                    fileStream.Value.CopyTo(memoryStream);
                                                    memoryStream.Position = 0;

                                                    try
                                                    {
                                                        var provider = new XMLStreamTemplateProvider();
                                                        ProvisioningTemplate template = provider.GetTemplate(memoryStream);

                                                        var applyingInformation = new ProvisioningTemplateApplyingInformation();
                                                        template.Parameters["BranchId"] = branch.branchId;
                                                        template.Parameters["BranchName"] = branch.BranchName;
                                                        siteContext.Web.ApplyProvisioningTemplate(template, applyingInformation);
                                                        _logger.LogWarning("Webpart added successfully.");
                                                    }
                                                    catch (ReflectionTypeLoadException ex)
                                                    {
                                                        foreach (var loaderEx in ex.LoaderExceptions)
                                                        {
                                                            _logger.LogError($"Webpart : {loaderEx.Message}");
                                                        }
                                                    }
                                                }
                                                UpdateBranchListItems(SiteURL, accessToken, branch.branchId, branchPlannerPlan, branchDocLib, siteUrl + "/SitePages/Home.aspx");
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger.LogError(ex.Message, "Failed to retrieve item details for ID: {ItemId}", changeItem.ItemId);
                                }
                                break;
                            case Microsoft.SharePoint.Client.ChangeType.Update:
                                _logger.LogWarning("Item was Updated.");
                                break;
                            case Microsoft.SharePoint.Client.ChangeType.DeleteObject:
                                _logger.LogWarning("Item was Deleted.");
                                break;
                        }
                        lastChangeToken = change.ChangeToken;
                    }
                }

                if (lastChangeToken != null)
                {
                    var updatedTokenEntity = new ChangeTokenEntity(partitionKey, rowKey)
                    {
                        ChangeToken = lastChangeToken.StringValue
                    };

                    var upsertOperation = TableOperation.InsertOrReplace(updatedTokenEntity);
                    await table.ExecuteAsync(upsertOperation);

                    _logger.LogWarning("Saved new ChangeToken: {Token}", lastChangeToken.StringValue);
                }
                else
                {
                    _logger.LogWarning("No changes detected...");
                }

            }
        }
        catch (Exception ex)
        {
            _logger.LogError("Error : ",ex.Message);
            throw;
        }
    }

    public void UpdateBranchListItems(string site, string accessToken, string branchId, string branchPlanner, string branchDocLibrary, string siteUrl)
    {
        _logger.LogWarning("Updating Branch List...");
        using (var context = new ClientContext(site))
        {
            context.ExecutingWebRequest += (sender, e) =>
            {
                e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
            };

            var branchList = context.Web.Lists.GetByTitle("Branch List");
            var branch = branchList.GetItems(CamlQuery.CreateAllItemsQuery());
            context.Load(branch);
            context.ExecuteQuery();

            foreach (var item in branch)
            {
                if ((string)item["BranchId"] == branchId)
                {
                    item["Siteurl"] = siteUrl;
                    item["LibraryUrl"] = branchDocLibrary;
                    item["Plannerurl"] = branchPlanner;

                    item.Update();
                    context.ExecuteQuery();

                    Console.WriteLine("Branch List updated: " + branchId);
                    break;
                }
            }
        }
        _logger.LogWarning("Branch List Updated...");
    }

    public List<(string BranchName, string branchId)> GetBranch(string site, string accessToken, string clientNumber)
    {
        _logger.LogWarning("Getting branches...");
        var branches = new List<(string, string)>();
        using (var context = new ClientContext(site))
        {
            context.ExecutingWebRequest += (sender, e) =>
            {
                e.WebRequestExecutor.WebRequest.Headers["Authorization"] = "Bearer " + accessToken;
            };

            var list = context.Web.Lists.GetByTitle("Branch List");
            CamlQuery query = CamlQuery.CreateAllItemsQuery();
            var items = list.GetItems(query);
            context.Load(items);
            try
            {
                context.ExecuteQuery();
            }
            catch (Exception ex)
            {
                _logger.LogError("Error while loading item : ",ex);
            }
            foreach (var item in items)
            {
                var clientField = item["Client"] as Microsoft.SharePoint.Client.FieldLookupValue;
                if (clientField != null && clientField.LookupValue.ToString() == clientNumber)
                {
                    var BranchName = item["BranchName"]?.ToString() ?? "";
                    var branchId = item["BranchId"]?.ToString() ?? "";
                    branches.Add((BranchName, branchId));
                }
            }
        }
        _logger.LogWarning("Return branches...");
        return branches;
    }

    public class ChangeTokenEntity : Microsoft.WindowsAzure.Storage.Table.TableEntity
    {
        public ChangeTokenEntity() { }

        public ChangeTokenEntity(string partitionKey, string rowKey)
        {
            PartitionKey = partitionKey;
            RowKey = rowKey;
        }

        public string ChangeToken { get; set; }
        public DateTime Timestamp { get; set; }
    }
}
