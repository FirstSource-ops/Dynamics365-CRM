using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using System;
using System.ServiceModel.Description;
using System.Configuration;
using Microsoft.Xrm.Sdk.Extensions;

namespace UtilityProject1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Starting Dynamics 365 Opportunity Close Example...");
                string clientId = "48a5b21c-0523-4f57-892d-e9339a831fae";
                string clientSecret = "qM28Q~6azAY1.Ovop_jso8aAhdbo3bzPOszlBds-";
                string authority = "https://login.microsoftonline.com/63a6ea12-4c41-4909-9a2d-e047fe21a297";
                string crmURL = "https://org3503feab.crm.dynamics.com/";
                string conString = $"AuthType=ClientSecret;" +
                    $"Url={crmURL};" +
                    $"Clientid={clientId};" +
                    $"ClientSecret={clientSecret};" +
                    $"Authority={authority};" +
                    $"RequirenNewInstance=True;";
                CrmServiceClient crmService = new CrmServiceClient(conString);

                if (crmService.IsReady)
                {
                    Console.WriteLine("Connected to Dynamics 365");


                    IOrganizationService orgService = (IOrganizationService)crmService.OrganizationWebProxyClient != null ? (IOrganizationService)crmService.OrganizationWebProxyClient : (IOrganizationService)crmService.OrganizationServiceProxy;


                    Guid opportunityId = new Guid("3cbbd39d-d3f0-ea11-a815-000d3a33f3c3");


                    var loseOpportunityRequest = new LoseOpportunityRequest
                    {
                        OpportunityClose = new Entity("opportunityclose")
                        {
                            Attributes =
                        {
                            { "opportunityid", new EntityReference("opportunity", opportunityId) },
                            { "subject", "Lost the Opportunity!" }
                        }
                        },

                        Status = new OptionSetValue(4)
                    };

                    var loseOpportunityResponse = (LoseOpportunityResponse)orgService.Execute(loseOpportunityRequest);

                    if (loseOpportunityResponse != null)
                    {
                        Console.WriteLine("Opportunity marked as lost successfully.");
                    }
                    else
                    {
                        Console.WriteLine("Failed to mark the opportunity as lost.");
                    }
                }
                else
                {
                    Console.WriteLine("Failed to connect to Dynamics 365. Check your connection string and credentials.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

    }
}
