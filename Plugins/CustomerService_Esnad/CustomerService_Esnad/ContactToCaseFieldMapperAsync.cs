using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace CustomerService_Esnad
{
    public class ContactToCaseFieldMapperAsync : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                tracingService.Trace("Plugin execution started.");

                if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity targetEntity))
                    return;

                if (targetEntity.LogicalName != "incident")
                    return;

                Guid caseId = targetEntity.Id;
                if (caseId == Guid.Empty)
                {
                    tracingService.Trace("Case ID is empty.");
                    return;
                }

                tracingService.Trace("Retrieving case with ID: " + caseId);
                Entity caseRecord = service.Retrieve("incident", caseId, new ColumnSet("customerid"));

                if (!caseRecord.Contains("customerid") || !(caseRecord["customerid"] is EntityReference customerRef))
                {
                    tracingService.Trace("Customer is missing or invalid.");
                    return;
                }

                Entity update = new Entity("incident", caseId);

                if (customerRef.LogicalName == "contact")
                {
                    tracingService.Trace("Customer is a contact. Retrieving contact details...");
                    Entity contact = service.Retrieve("contact", customerRef.Id, new ColumnSet("fullname", "emailaddress1", "mobilephone", "new_nationalidnumber", "new_companyname"));

                    string companyName = string.Empty;
                    string crNumber= string.Empty;

                    if (contact.Contains("new_companyname") && contact["new_companyname"] is EntityReference companyRef)
                    {
                        tracingService.Trace("Contact has a company lookup. Retrieving account...");

                        // Retrieve both name and CR number from the account
                        Entity account = service.Retrieve("account", companyRef.Id, new ColumnSet("name", "new_crnumber"));

                        companyName = account.GetAttributeValue<string>("name") ?? string.Empty;

                        crNumber = account.GetAttributeValue<string>("new_crnumber") ?? string.Empty;
                        tracingService.Trace("Retrieved company name: " + companyName);
                        tracingService.Trace("Retrieved CR number: " + crNumber);

                        
                    }


                    update["new_customername"] = contact.GetAttributeValue<string>("fullname") ?? string.Empty;
                    update["new_email"] = contact.GetAttributeValue<string>("emailaddress1") ?? string.Empty;
                    update["new_phonenumber"] = contact.GetAttributeValue<string>("mobilephone") ?? string.Empty;
                    update["new_nationalidnumber"] = contact.GetAttributeValue<string>("new_nationalidnumber") ?? string.Empty;
                    update["new_companeyname"] = companyName;
                    // Set CR number on the Case record
                    update["new_crnumber"] = crNumber;
                }
                else if (customerRef.LogicalName == "account")
                {
                    tracingService.Trace("Customer is an account. Retrieving account details...");
                    Entity account = service.Retrieve("account", customerRef.Id, new ColumnSet("name", "emailaddress1", "new_companyrepresentativephonenumber", "new_crnumber"));

                    update["new_email"] = account.GetAttributeValue<string>("emailaddress1") ?? string.Empty;
                    update["new_phonenumber"] = account.GetAttributeValue<string>("new_companyrepresentativephonenumber") ?? string.Empty;
                    update["new_companeyname"] = account.GetAttributeValue<string>("name") ?? string.Empty;
                    update["new_crnumber"] = account.GetAttributeValue<string>("new_crnumber") ?? string.Empty;
                }
                else
                {
                    tracingService.Trace("Customer is neither contact nor account.");
                    return;
                }

                tracingService.Trace("Updating case with mapped fields...");
                service.Update(update);
                tracingService.Trace("Case updated successfully.");
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("❌ Contact-to-Case plugin failed: " + ex.Message, ex);
            }
        }
    }
}
