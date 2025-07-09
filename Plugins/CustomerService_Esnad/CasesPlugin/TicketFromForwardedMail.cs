using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace CasesPlugin
{
    public class TicketFromForwardedMail : IPlugin
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
                    if (entity.Attributes.Contains("directioncode")) // 1 = Incoming
                    {
                        // Check "from" field
                        if (entity.Attributes.Contains("from"))
                        {
                            EntityCollection fromCollection = entity.GetAttributeValue<EntityCollection>("from");
                            if (fromCollection != null && fromCollection.Entities.Count > 0)
                            {
                                foreach (var activityParty in fromCollection.Entities)
                                {
                                    Entity party = activityParty;
                                    string senderEmail = party.GetAttributeValue<string>("addressused");

                                    if (!string.IsNullOrEmpty(senderEmail) &&
                                        senderEmail.Equals("no-reply@taadeen.sa", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string plainTextBody = entity.GetAttributeValue<string>("description") ?? string.Empty;

                                        string body = StripHtml(plainTextBody);

                                        string subject1 = entity.GetAttributeValue<string>("subject");
                                        if (subject1.ToLower() == "book an appointment")
                                        //           if (!istrue)
                                        {
                                            EntityReference customerRef = null;
                                            //string beneficiaryType = MatchValueForAppointment(body, @"Type of Beneficiary\s*\n(.+)");
                                            //string idNumber = MatchValueForAppointment(body, @"ID Number\s*\n(.+)");
                                            //string fullName = MatchValueForAppointment(body, @"Full Name\s*\n(.+)");
                                            //string phone2 = MatchValueForAppointment(body, @"Phone Number\s*\n(.+)");
                                            //string email1 = MatchValueForAppointment(body, @"Email Address\s*\n(.+)");
                                            //string requestType2 = MatchValueForAppointment(body, @"Request Type\s*\n(.+)");
                                            //string department = MatchValueForAppointment(body, @"Sector\s*\n(.+)");
                                            //string reason = MatchValueForAppointment(body, @"Reason of this request\s*\n(.+)");

                                            string beneficiaryType = MatchValueForAppointment(body, @"(?:Type of Beneficiary|مقدم الطلب)\s*\n(.+)");
                                            string idNumber = MatchValueForAppointment(body, @"(?:ID Number|رقم الهوية)\s*\n(.+)");

                                            string companyName = MatchValueForAppointment(body, @"(?:Company Name|اسم الشركة)\s*\n(.+)");
                                            string CRN = MatchValueForAppointment(body, @"(?:Commercial Registration Number|رقم السجل التجاري)\s*\n(.+)");

                                            string fullName = MatchValueForAppointment(body, @"(?:Full Name|الإسم الثلاثي)\s*\n(.+)");
                                            string phone2 = MatchValueForAppointment(body, @"(?:Phone Number|رقم الهاتف)\s*\n(.+)").Replace("&#43;", "+");
                                            string email1 = MatchValueForAppointment(body, @"(?:Email Address|البريد الإلكتروني)\s*\n(.+)");
                                            string requestType2 = MatchValueForAppointment(body, @"(?:Request Type|نوع الموعد)\s*\n(.+)");
                                            string department = MatchValueForAppointment(body, @"(?:Sector|القطاع)\s*\n(.+)");
                                            // string complianceType = MatchValueForAppointment(body, @"(?:Compliance|الامتثال)\s*\n(.+)");
                                            string complianceType = MatchValueForAppointment(
                                                         body,
                                                         @"(?<=\r?\n\r?\n)(?:الامتثال|Compliance)\s*\r?\n\s*([^\r\n]+)"
                                                     );
                                            //string licenseType = MatchValueForAppointment(body, @"(?:Licenses|الرخص)\s*\n(.+)");                                                                                                          //  string licenseType = MatchValue(body, @"License Type)[^\r\n:]*[:\-]?\s*([^\(]+)").Trim();
                                            //string licenseType = MatchValueForAppointment(body, @"(?:الرخص|License Type)\s*\r?\n\s*([^\r\n]+)");
                                            string licenseType = MatchValueForAppointment(
                                                        body,
                                                        @"(?<=\r?\n\r?\n)(?:الرخص|Licenses)\s*\r?\n\s*([^\r\n]+)"
                                                    );

                                            // string licenseType = MatchValueForAppointment(body, @"(?:License Type|الرخص)\s*\n(.+)");
                                            string reason = MatchValueForAppointment(body, @"(?:Reason of this request|سبب حجز الموعد)\s*\n([\s\S]+)"); // Supports multiline







                                            //var normalizedType = NormalizeAppointmentType(requestType2);
                                            //var normalizedSector = NormalizeSector(department);



                                            if (beneficiaryType == "فرد" || beneficiaryType.ToLower() == "individual")
                                            {
                                                var query = new QueryExpression("contact")
                                                {
                                                    ColumnSet = new ColumnSet("contactid"),
                                                    Criteria = new FilterExpression
                                                    {
                                                        Conditions = {
                                new ConditionExpression("emailaddress1", ConditionOperator.Equal, email1)
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
                                                        ["emailaddress1"] = email1,
                                                        ["mobilephone"] = phone2,
                                                        ["new_nationalidnumber"] = idNumber,
                                                    };
                                                    contactId = service.Create(contact);
                                                }

                                                customerRef = new EntityReference("contact", contactId);
                                            }
                                            //investor
                                            else if (beneficiaryType == "شركة" || beneficiaryType.ToLower() == "company")
                                            {
                                                var query = new QueryExpression("account")
                                                {
                                                    ColumnSet = new ColumnSet("accountid"),
                                                    Criteria = new FilterExpression
                                                    {
                                                        Conditions = {
                                new ConditionExpression("name", ConditionOperator.Equal, companyName)
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
                                                        ["name"] = companyName,
                                                        ["emailaddress1"] = email1,
                                                        ["new_companyrepresentativephonenumber"] = phone2,
                                                        ["new_crnumber"] = CRN,
                                                        ["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("70FA9BC3-6D4B-F011-A3FE-D4DE6FAB9C57"))
                                                    };
                                                    accountId = service.Create(account);
                                                }

                                                customerRef = new EntityReference("account", accountId);
                                            }
                                            else
                                            {
                                                Console.WriteLine("❌ Unsupported beneficiary type: " + beneficiaryType);
                                                return;
                                            }

                                            var incident = new Entity("incident")
                                            {
                                                ["title"] = "book an appointment",
                                                ["description"] = reason,
                                                //["new_requesttype"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_requesttype", requestType2)),

                                                ["new_beneficiarytype"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_beneficiarytype", ConvertBeneficiaryTypeToEnglish(beneficiaryType))),
                                                ["new_ticketsubmissionchannel"] = new OptionSetValue(4), // Email

                                                //["new_compliance"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_compliance", complianceType)),
                                                ["customerid"] = customerRef,
                                                ["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("70FA9BC3-6D4B-F011-A3FE-D4DE6FAB9C57"))

                                            };
                                            //string normalizedLicense = NormalizeLicenseType(licenseType);
                                            //if (!string.IsNullOrEmpty(normalizedLicense))
                                            //{
                                            //    incident["new_licenses"] = new OptionSetValue(
                                            //        GetOptionSetValue(service, "incident", "new_licenses", normalizedLicense)
                                            //    );
                                            //}
                                            //string normalizedCompliance = NormalizeComplianceType(complianceType);

                                            //if (!string.IsNullOrWhiteSpace(normalizedCompliance))
                                            //{
                                            //    // Only try to set OptionSet if a valid value is found
                                            //    try
                                            //    {
                                            //        incident["new_compliance"] = new OptionSetValue(
                                            //            GetOptionSetValue(service, "incident", "new_compliance", normalizedCompliance)
                                            //        );
                                            //    }
                                            //    catch
                                            //    {
                                            //        // Silently skip if OptionSet value is not found in metadata
                                            //        // Or optionally log to a Note or Trace
                                            //    }
                                            //}
                                            // ➕ Set new_licensetype if sector is Licensing
                                            if ((department.Equals("Licensing", StringComparison.OrdinalIgnoreCase) ||
                                                 department.Equals("الرخص", StringComparison.OrdinalIgnoreCase)) &&
                                                !string.IsNullOrWhiteSpace(licenseType))
                                            {

                                                incident["new_licenses"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_licenses", licenseType));
                                            }
                                            // ➕ Set new_compliancetype if sector is Compliance
                                            else if ((department.Equals("Compliance", StringComparison.OrdinalIgnoreCase) ||
                                                      department.Equals("الامتثال", StringComparison.OrdinalIgnoreCase)) &&
                                                     !string.IsNullOrWhiteSpace(complianceType))
                                            {
                                                incident["new_compliance"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_compliance", complianceType));

                                            }


                                            incident["new_sector"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_sector", department));
                                            Entity email11 = new Entity("email");
                                            email11.Id = entity.Id;
                                            email11.Attributes["regardingobjectid"] = new EntityReference("incident", service.Create(incident));
                                            service.Update(email11);

                                        }


                                        else if (subject1.ToLower() == "contact us")
                                        {
                                            string beneficiaryType = MatchValue(body, @"(?:نوع المستفيد|Type of Beneficiary)[:\-]?\s*(مستثمر|فرد|Investor|Individual)");
                                            string company = MatchValue(body, @"(?:اسم الشركة|Company Name)[:\-]?\s*([^\r\n]+?)\s*(?=رقم السجل التجاري|Commercial Registration Number)");
                                            string crNumber = MatchValue(body, @"(?:رقم السجل التجاري|Commercial Registration Number(?:\s*\(CR\))?)[:\-]?\s*(\d{5,})");

                                            //string crNumber = MatchValue(body, @"(?:رقم السجل التجاري|Commercial Registration Number.*CR.*)[:\-]?\s*(\d+)");
                                            string phone = MatchValue(body, @"(?:رقم الهاتف|Mobile Number)[:\-]?\s*(\d+)").Replace("&#43;", "+");
                                            string emailAddr = MatchValue(body, @"(?:عنوان البريد الإلكتروني|Email Address)[:\-]?\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})");
                                            string requestType = MatchValue(body, @"(?:نوع الطلب|Request Type)[:\-]?\s*(.+?)\s*(?=الموضوع|Subject)");
                                            string subject = MatchValue(body, @"(?:الموضوع|Subject)[:\-]?\s*(.+?)\s*(?=نص الرسالة|Message Text)");
                                            string message = MatchValue(body, @"(?:نص الرسالة|Message Text)[:\-]?\s*([\s\S]+?)(?=تحميل ملفات|Attachments|$)");
                                            string attachment = MatchValue(body, @"(?:تحميل ملفات|Attachments)[:\-]?\s*([^\r\n\(]+)");
                                            string nationalId = MatchValue(body, @"(?:رقم الهوية|National ID Number)[:\-]?\s*(\d{10,15})");





                                            EntityReference customerRef;
                                            //individual
                                            if (beneficiaryType == "فرد" || beneficiaryType.ToLower() == "individual")
                                            {
                                                var query = new QueryExpression("contact")
                                                {
                                                    ColumnSet = new ColumnSet("contactid"),
                                                    Criteria = new FilterExpression
                                                    {
                                                        Conditions = {
                                new ConditionExpression("emailaddress1", ConditionOperator.Equal, emailAddr)
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
                                                        ["lastname"] = emailAddr,
                                                        ["emailaddress1"] = emailAddr,
                                                        ["mobilephone"] = phone,
                                                        ["new_nationalidnumber"] = nationalId,
                                                        // ["new_companyname"] = GetOrCreateCompany(service, company)
                                                    };
                                                    contactId = service.Create(contact);
                                                }

                                                customerRef = new EntityReference("contact", contactId);
                                            }
                                            //investor
                                            else if (beneficiaryType == "مستثمر" || beneficiaryType.ToLower() == "investor")
                                            {
                                                var query = new QueryExpression("account")
                                                {
                                                    ColumnSet = new ColumnSet("accountid"),
                                                    Criteria = new FilterExpression
                                                    {
                                                        Conditions = {
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
                                                        ["new_crnumber"] = crNumber,
                                                        ["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("70FA9BC3-6D4B-F011-A3FE-D4DE6FAB9C57"))
                                                    };
                                                    accountId = service.Create(account);
                                                }

                                                customerRef = new EntityReference("account", accountId);
                                            }
                                            else
                                            {
                                                return;
                                            }

                                            var incident = new Entity("incident")
                                            {
                                                ["title"] = subject,
                                                ["description"] = message,
                                                ["new_tickettype"] = GetLookupByName(service, "new_tickettype", requestType),

                                                ["new_beneficiarytype"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_beneficiarytype", ConvertBeneficiaryTypeToEnglish(beneficiaryType))),
                                                ["new_ticketsubmissionchannel"] = new OptionSetValue(100000000), // Email
                                                ["customerid"] = customerRef,
                                                ["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("70FA9BC3-6D4B-F011-A3FE-D4DE6FAB9C57"))
                                            };

                                            Entity email11 = new Entity("email");
                                            email11.Id = entity.Id;
                                            email11.Attributes["regardingobjectid"] = new EntityReference("incident", service.Create(incident));
                                            service.Update(email11);

                                        }
                                    }
                                }
                            }
                        }
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
            }
            catch (Exception ex)
            {
                tracer.Trace("❌ Plugin Exception: " + ex.ToString());
                throw;
            }
        }
        static string MatchValueForAppointment(string input, string pattern)
        {
            var match = Regex.Match(input, pattern, RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value.Trim() : "Not Found";
        }

        public static string StripHtml(string html)
        {
            try
            {
                var doc = System.Xml.Linq.XDocument.Parse($"<root>{html}</root>");
                return string.Concat(doc.DescendantNodes().OfType<System.Xml.Linq.XText>().Select(t => t.Value));
            }
            catch
            {
                // fallback if parsing fails
                return Regex.Replace(html, "<.*?>", string.Empty);
            }
        }
        //private static string StripHtml(string html)
        //{
        //    var doc = new HtmlAgilityPack.HtmlDocument();
        //    doc.LoadHtml(html);
        //    return HtmlEntity.DeEntitize(doc.DocumentNode.InnerText);
        //}
        private static string ConvertBeneficiaryTypeToEnglish(string arabicType)
        {
            switch (arabicType.Trim())
            {
                case "مستثمر":
                    return "Investor";
                case "Investor":
                    return "Investor";

                case "فرد":
                    return "Individual";
                case "Individual":
                    return "Individual";
                default:
                    return "Unknown"; // Or return arabicType if you want to keep unrecognized values
            }
        }

        private static string MatchValue(string input, string pattern)
        {
            input = System.Net.WebUtility.HtmlDecode(input); // decode HTML entities
            input = input.Replace("\u00A0", " "); // non-breaking spaces
            input = Regex.Replace(input, @"\s+", " "); // normalize all whitespace

            var match = Regex.Match(input, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
            return match.Success ? match.Groups[1].Value.Trim() : string.Empty;
        }
        private static EntityReference GetOrCreateCompany(IOrganizationService service, string companyName)
        {
            var query = new QueryExpression("account")
            {
                ColumnSet = new ColumnSet("accountid"),
                Criteria = new FilterExpression
                {
                    Conditions = {
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

        private static EntityReference GetLookupByName(IOrganizationService service, string entityName, string name)
        {
            var query = new QueryExpression(entityName)
            {
                ColumnSet = new ColumnSet($"{entityName}id"),
                Criteria = new FilterExpression
                {
                    Conditions = {
                        new ConditionExpression("new_tickettype", ConditionOperator.Equal, name)
                    }
                }
            };

            var result = service.RetrieveMultiple(query);
            if (result.Entities.Count > 0)
                return new EntityReference(entityName, result.Entities[0].Id);

            throw new InvalidPluginExecutionException($"Lookup '{name}' not found in {entityName}.");
        }

        private static int GetOptionSetValue(IOrganizationService service, string entityName, string fieldName, string label)
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
