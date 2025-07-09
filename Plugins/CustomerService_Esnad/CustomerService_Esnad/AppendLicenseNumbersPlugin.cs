using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;

namespace CustomerService_Esnad
{

    public class AppendLicenseNumbersPlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference.
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            // Check if the context is for an update operation and the correct entity is being updated.
            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity caseEntity)
            {
                if (caseEntity.LogicalName != "incident")
                {
                    throw new InvalidPluginExecutionException("This plugin only works with incident entity.");
                }

                // Retrieve the customer (could be an account or contact)
                EntityReference customerReference = caseEntity.GetAttributeValue<EntityReference>("customerid");

                if (customerReference == null)
                {
                    throw new InvalidPluginExecutionException("This case is not linked to any customer.");
                }

                // Check if the customer is an account or a contact and process accordingly
                Guid customerId = customerReference.Id;
                string customerLogicalName = customerReference.LogicalName;

                if (customerLogicalName == "account")
                {
                    AppendLicensesToAccount(service, customerId, caseEntity.Id);
                }
                else if (customerLogicalName == "contact")
                {
                    AppendLicensesToContact(service, customerId, caseEntity.Id);
                }
                else
                {
                    throw new InvalidPluginExecutionException("This case is linked to an unsupported customer type.");
                }
            }
        }

        private void AppendLicensesToAccount(IOrganizationService service, Guid accountId, Guid caseId)
        {
            // Fetch all licenses (new_licensetype) linked to the account (subgrid of Account)
            var licenseQuery = new QueryExpression("new_licensetype")
            {
                ColumnSet = new ColumnSet("new_licensenumber"),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression("new_companyname", ConditionOperator.Equal, accountId)
                }
                }
            };

            // Execute the query to fetch all licenses associated with the account
            var licenses = service.RetrieveMultiple(licenseQuery).Entities;

            // Create a list to hold the license numbers
            var licenseNumbers = new List<string>();

            // Iterate through the licenses and add their numbers
            foreach (var license in licenses)
            {
                string licenseNumber = license.GetAttributeValue<string>("new_licensenumber");
                if (!string.IsNullOrEmpty(licenseNumber))
                {
                    licenseNumbers.Add(licenseNumber);
                }
            }

            // Combine the license numbers into a comma-separated string
            string licenseNumbersCsv = string.Join(",", licenseNumbers);

            // Update the case with the license numbers
            UpdateCaseWithLicenses(service, licenseNumbersCsv, caseId);
        }

        private void AppendLicensesToContact(IOrganizationService service, Guid contactId, Guid caseId)
        {
            // Fetch the account associated with the contact (new_companyname)
            Entity contact = service.Retrieve("contact", contactId, new ColumnSet("new_companyname"));
            EntityReference accountReference = contact.GetAttributeValue<EntityReference>("new_companyname");
            if (accountReference == null)
            {
                throw new InvalidPluginExecutionException("This contact is not linked to an account.");
            }

            Guid accountId = accountReference.Id;

            // Fetch all licenses (new_licensetype) linked to the account (subgrid of Account)
            var licenseQuery = new QueryExpression("new_licensetype")
            {
                ColumnSet = new ColumnSet("new_licensenumber"),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression("new_companyname", ConditionOperator.Equal, accountId)
                }
                }
            };

            // Execute the query to fetch all licenses associated with the account
            var licenses = service.RetrieveMultiple(licenseQuery).Entities;

            // Create a list to hold the license numbers
            var licenseNumbers = new List<string>();

            // Iterate through the licenses and add their numbers
            foreach (var license in licenses)
            {
                string licenseNumber = license.GetAttributeValue<string>("new_licensenumber");
                if (!string.IsNullOrEmpty(licenseNumber))
                {
                    licenseNumbers.Add(licenseNumber);
                }
            }

            // Combine the license numbers into a comma-separated string
            string licenseNumbersCsv = string.Join(",", licenseNumbers);

            // Update the case with the license numbers
            UpdateCaseWithLicenses(service, licenseNumbersCsv, caseId);
        }

        private void UpdateCaseWithLicenses(IOrganizationService service, string licenseNumbersCsv, Guid caseId)
        {
            // Update the case with the license numbers
            Entity updateCase = new Entity("incident", caseId)
            {
                ["new_globalsearch"] = licenseNumbersCsv
            };

            // Update the case record with the new value
            service.Update(updateCase);
        }
    }
}
