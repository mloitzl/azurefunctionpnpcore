using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using System.Linq;
using PnP.Core;
using PnP.Core.Model.SharePoint;

namespace mloitzl.sharepoint.functions
{

    public class CreateTerm
    {
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly ILogger<CreateTerm> _logger;

        public CreateTerm(IPnPContextFactory pnpContextFactory, ILogger<CreateTerm> logger)
        {
            _pnpContextFactory = pnpContextFactory;
            _logger = logger;
        }

        [FunctionName("CreateTerm")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = "term/{name}")] HttpRequest req, string name)
        {
            using (var pnpContext = await _pnpContextFactory.CreateAsync("Default"))
            {
                _logger.LogInformation("Creating Term '{DefaultLabel}'", name);
                try
                {
                    ITermGroup termGroup = pnpContext.TermStore.Groups.First();
                    ITermSet termSet = termGroup.Sets.First();

                    var newTerm = await termSet.Terms.AddAsync(name);
                    return new OkObjectResult(new { id = newTerm.Id });
                }
                catch (MicrosoftGraphServiceException ex)
                {
                    _logger.LogError((Exception)ex, ex.Error.ToString());
                    return new BadRequestObjectResult(ex.Error);
                }
            }
        }
    }
}
