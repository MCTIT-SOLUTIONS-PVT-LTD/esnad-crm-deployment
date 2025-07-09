using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
namespace CustomerService_Esnad
{
    public class CaseOnForwardedEmail : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            ITracingService tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            try
            {
                //if (context.MessageName != "Create" || !context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity email))
                //    return;
                if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity entity)
                {

                    string body = entity.GetAttributeValue<string>("description") ?? string.Empty;

                    tracer.Trace("✅ Extracting fields from email body...");

                    // Use regex to extract fields from body (supports Arabic/English)
                    string company = MatchValue(body, @"(?<=اسم الشركة\s*)[^\r\n]+");
                    string crNumber = MatchValue(body, @"(?<=رقم السجل التجاري\s*)\d+");
                    string fullName = MatchValue(body, @"(?<=اسم المستفيد\s*)[^\r\n]+");
                    string nationalId = MatchValue(body, @"(?<=رقم الهوية\s*)\d+");
                    string phone = MatchValue(body, @"(?<=رقم الهاتف\s*)\d+");
                    string emailAddr = MatchValue(body, @"(?<=عنوان البريد الإلكتروني\s*)\S+");
                    string beneficiaryType = MatchValue(body, @"(?<=نوع المستفيد\s*)[^\r\n]+");
                    string requestType = MatchValue(body, @"(?<=نوع الطلب\s*)[^\r\n]+");
                    string subject = MatchValue(body, @"(?<=الموضوع\s*)[^\r\n]+");
                    string message = MatchValue(body, @"(?<=نص الرسالة\s*)[\s\S]+"); // multi-line support

                    EntityReference customerRef;

                    // Create or find Contact
                    if (beneficiaryType == "فرد")
                    {
                        var query = new QueryExpression("contact")
                        {
                            ColumnSet = new ColumnSet("contactid"),
                            Criteria = new FilterExpression
                            {
                                Conditions =
                        {
                            new ConditionExpression("new_nationalidnumber", ConditionOperator.Equal, nationalId)
                        }
                            }
                        };

                        var result = service.RetrieveMultiple(query);
                        Guid contactId;

                        if (result.Entities.Count > 0)
                        {
                            contactId = result.Entities[0].Id;
                        }
                        else
                        {
                            var contact = new Entity("contact")
                            {
                                ["lastname"] = fullName,
                                ["emailaddress1"] = emailAddr,
                                ["mobilephone"] = phone,
                                ["new_nationalidnumber"] = nationalId,
                                ["new_companyname"] = GetOrCreateCompany(service, company)
                            };
                            contactId = service.Create(contact);
                        }

                        customerRef = new EntityReference("contact", contactId);
                    }
                    // Create or find Account
                    else if (beneficiaryType == "مستثمر")
                    {
                        var query = new QueryExpression("account")
                        {
                            ColumnSet = new ColumnSet("accountid"),
                            Criteria = new FilterExpression
                            {
                                Conditions =
                        {
                            new ConditionExpression("name", ConditionOperator.Equal, company)
                        }
                            }
                        };

                        var result = service.RetrieveMultiple(query);
                        Guid accountId;

                        if (result.Entities.Count > 0)
                        {
                            accountId = result.Entities[0].Id;
                        }
                        else
                        {
                            var account = new Entity("account")
                            {
                                ["name"] = company,
                                ["emailaddress1"] = emailAddr,
                                ["new_companyrepresentativephonenumber"] = phone,
                                ["new_crnumber"] = crNumber
                            };
                            accountId = service.Create(account);
                        }

                        customerRef = new EntityReference("account", accountId);
                    }
                    else
                    {
                        tracer.Trace("❌ Unsupported beneficiary type: " + beneficiaryType);
                        return;
                    }

                    // Create Case (incident)
                    var incident = new Entity("incident")
                    {
                        ["title"] = subject,
                        ["description"] = message,
                        ["new_tickettype"] = GetLookupByName(service, "new_tickettype", requestType),
                        ["new_beneficiarytype"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_beneficiarytype", beneficiaryType)),
                        ["new_ticketsubmissionchannel"] = new OptionSetValue(100000000), // Email
                        ["customerid"] = customerRef
                    };

                    service.Create(incident);
                }
                // Safely working with the Target entity
                //string logicalName = entity.LogicalName;
                //Entity email = context.InputParameters["Target"];
                //// Only process incoming emails
                //if (!email.GetAttributeValue<bool>("directioncode"))
                //    return;

                //// Extract sender email
                //if (!email.Attributes.Contains("from"))
                //    return;

                //var fromParties = email.GetAttributeValue<EntityCollection>("from");
                //if (fromParties.Entities.Count == 0)
                //    return;

                //var senderParty = fromParties.Entities[0];
                //var partyIdRef = senderParty.GetAttributeValue<EntityReference>("partyid");
                //if (partyIdRef == null || !partyIdRef.Name.Equals("no-reply@taadeen.sa", StringComparison.OrdinalIgnoreCase))
                //    return;


            }
            catch (Exception ex)
            {
                tracer.Trace("❌ Plugin Exception: " + ex.ToString());
                throw;
            }
        }

        private static string MatchValue(string input, string pattern)
        {
            var match = Regex.Match(input, pattern, RegexOptions.IgnoreCase);
            return match.Success ? match.Value.Trim() : string.Empty;
        }

        private EntityReference GetOrCreateCompany(IOrganizationService service, string companyName)
        {
            var query = new QueryExpression("account")
            {
                ColumnSet = new ColumnSet("accountid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression("name", ConditionOperator.Equal, companyName)
                }
                }
            };

            var result = service.RetrieveMultiple(query);
            if (result.Entities.Count > 0)
                return new EntityReference("account", result.Entities[0].Id);

            var account = new Entity("account") { ["name"] = companyName };
            var id = service.Create(account);
            return new EntityReference("account", id);
        }

        private EntityReference GetLookupByName(IOrganizationService service, string entityName, string name)
        {
            var query = new QueryExpression(entityName)
            {
                ColumnSet = new ColumnSet($"{entityName}id"),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression("name", ConditionOperator.Equal, name)
                }
                }
            };

            var result = service.RetrieveMultiple(query);
            if (result.Entities.Count > 0)
                return new EntityReference(entityName, result.Entities[0].Id);

            throw new InvalidPluginExecutionException($"Lookup '{name}' not found in {entityName}.");
        }

        private int GetOptionSetValue(IOrganizationService service, string entityName, string fieldName, string label)
        {
            var response = (RetrieveAttributeResponse)service.Execute(new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = fieldName,
                RetrieveAsIfPublished = true
            });

            var metadata = (PicklistAttributeMetadata)response.AttributeMetadata;
            foreach (var opt in metadata.OptionSet.Options)
            {
                if (opt.Label.UserLocalizedLabel.Label == label)
                    return opt.Value.Value;
            }

            throw new InvalidPluginExecutionException($"Option '{label}' not found in '{fieldName}' on '{entityName}'");
        }
    }
}
