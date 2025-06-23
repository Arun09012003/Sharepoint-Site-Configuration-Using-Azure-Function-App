using Azure.Storage.Queues;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using System.Text;

namespace HelloAzureFunctions;

public class httpTrigger
{
    private readonly ILogger<httpTrigger> _logger;

    public httpTrigger(ILogger<httpTrigger> logger)
    {
        _logger = logger;
    }

    [Function("httpTrigger")]
    public async Task<IActionResult> RunAsync([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
        string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
        var queueClient = new QueueClient(Environment.GetEnvironmentVariable("AzureWebJobsStorage"), "myqueue-items");
        await queueClient.CreateIfNotExistsAsync();
        await queueClient.SendMessageAsync(Convert.ToBase64String(Encoding.UTF8.GetBytes(requestBody)));
        _logger.LogWarning("Webhook handled. New items queued.");
        return new OkObjectResult("Webhook handled. New items queued.");

    }
}