using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml;

namespace CustomerService_Esnad
{
    public class EmailToCasePlugin : IPlugin
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
                    tracer.Trace("📄 Email entity detected.");
                    if (entity.Attributes.Contains("directioncode")) // 1 = Incoming
                    {
                        tracer.Trace("📨 Email is incoming.");
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
                                    tracer.Trace($"📧 Sender email: {senderEmail}");
                                    if (!string.IsNullOrEmpty(senderEmail) &&
                                            senderEmail.Equals("no-reply@taadeen.sa", StringComparison.OrdinalIgnoreCase))
                                    {
                                        string plainTextBody = entity.GetAttributeValue<string>("description") ?? string.Empty;

                                        string body = StripHtml(plainTextBody);

                                        string beneficiaryType = MatchValue(body, @"(?:نوع المستفيد|Type of Beneficiary)[:\-]?\s*(مستثمر|فرد|Investor|Individual)");
                                        string company = MatchValue(body, @"(?:اسم الشركة|Company Name)[:\-]?\s*([^\r\n]+?)\s*(?=رقم السجل التجاري|Commercial Registration Number)");
                                        string crNumber = MatchValue(body, @"(?:رقم السجل التجاري|Commercial Registration Number(?:\s*\(CR\))?)[:\-]?\s*(\d{5,})");

                                        //string crNumber = MatchValue(body, @"(?:رقم السجل التجاري|Commercial Registration Number.*CR.*)[:\-]?\s*(\d+)");
                                        string phone = MatchValue(body, @"(?:رقم الهاتف|Mobile Number)[:\-]?\s*(\d+)");
                                        string emailAddr = MatchValue(body, @"(?:عنوان البريد الإلكتروني|Email Address)[:\-]?\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})");
                                        string requestType = MatchValue(body, @"(?:نوع الطلب|Request Type)[:\-]?\s*(.+?)\s*(?=الموضوع|Subject)");
                                        string subject = MatchValue(body, @"(?:الموضوع|Subject)[:\-]?\s*(.+?)\s*(?=نص الرسالة|Message Text)");
                                        string message = MatchValue(body, @"(?:نص الرسالة|Message Text)[:\-]?\s*([\s\S]+?)(?=تحميل ملفات|Attachments|$)");
                                        string attachment = MatchValue(body, @"(?:تحميل ملفات|Attachments)[:\-]?\s*([^\r\n\(]+)");
                                        string nationalId = MatchValue(body, @"(?:رقم الهوية|National ID Number)[:\-]?\s*(\d{10,15})");

                                        string appointmentType = MatchValue(body, @"(?:(?:نوع الموعد|Appointment Type)\s*[:\-]?\s*)([^\r\n]+)");
                                        var normalizedType = NormalizeAppointmentType(appointmentType);
                                        Console.WriteLine($"🔁 Normalized Appointment Type: {appointmentType} → {normalizedType}");

                                        string department = MatchValue(body, @"(?:(?:القطاع|Department|Sector)\s*[:\-]?\s*)([^\r\n]+)");
                                        var normalizedSector = NormalizeSector(department);
                                        Console.WriteLine($"🔁 Normalized Sector: {department} → {normalizedSector}");

                                        string reason = MatchValue(body, @"(?:(?:سبب حجز الموعد|Reason for Appointment)\s*[:\-]?\s*)([\s\S]+?)\r?\n");

                                        string licenseType = MatchValue(body, @"(?:الرخص|License Type)[^\r\n:]*[:\-]?\s*([^\(]+)").Trim();
                                        var normalizedLicenseType = NormalizeLicenseType(licenseType);
                                        Console.WriteLine($"🔁 Normalized License Type: {licenseType} → {normalizedLicenseType}");
                                        string complianceType = MatchValue(body, @"(?:الامتثال|Compliance)[^\r\n:]*[:\-]?\s*([^\(]+)").Trim();
                                        tracer.Trace("🔍 Extracted values:\n" +
                                        $"- BeneficiaryType: {beneficiaryType}\n" +
                                        $"- Company: {company}\n" +
                                        $"- CR Number: {crNumber}\n" +
                                        $"- Phone: {phone}\n" +
                                        $"- Email: {emailAddr}\n" +
                                        $"- Request Type: {requestType}\n" +
                                        $"- Subject: {subject}\n" +
                                        $"- Department: {department}\n" +
                                        $"- License Type: {licenseType}\n" +
                                        $"- Compliance Type: {complianceType}");


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
                                                    // ["new_nationalidnumber"] = nationalId,
                                                    ["new_companyname"] = GetOrCreateCompany(service, company)
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
                                            Console.WriteLine("❌ Unsupported beneficiary type: " + beneficiaryType);
                                            return;
                                        }

                                        // Set default values --vrushti
                                        if (string.IsNullOrWhiteSpace(subject))
                                            subject = "Book an appointment";

                                        if (string.IsNullOrWhiteSpace(requestType))
                                            requestType = "Suggestion";

                                        if (string.IsNullOrWhiteSpace(message))
                                            message = reason ?? "No message provided";

                                        var incident = new Entity("incident")
                                        {
                                            ["title"] = subject,
                                            ["description"] = message,
                                            ["new_tickettype"] = GetLookupByName(service, "new_tickettype", requestType),

                                            ["new_beneficiarytype"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_beneficiarytype", ConvertBeneficiaryTypeToEnglish(beneficiaryType))),
                                            ["new_ticketsubmissionchannel"] = new OptionSetValue(4), // Email
                                            ["customerid"] = customerRef,
                                            ["transactioncurrencyid"] = new EntityReference("transactioncurrency", new Guid("70FA9BC3-6D4B-F011-A3FE-D4DE6FAB9C57"))
                                        };
                                        // Optional: Set Appointment Type (new_requesttype OptionSet)
                                        if (!string.IsNullOrWhiteSpace(appointmentType))
                                        {
                                            incident["new_requesttype"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_requesttype", normalizedType));
                                        }

                                        // Optional: Set Department / Sector (new_sector OptionSet)
                                        if (!string.IsNullOrWhiteSpace(department))
                                        {
                                            // Set the sector (new_sector OptionSet)
                                            incident["new_sector"] = new OptionSetValue(GetOptionSetValue(service, "incident", "new_sector", normalizedSector));

                                            // ➕ Set new_licensetype if sector is Licensing
                                            if ((department.Equals("Licensing", StringComparison.OrdinalIgnoreCase) ||
                                                 department.Equals("الرخص", StringComparison.OrdinalIgnoreCase)) &&
                                                !string.IsNullOrWhiteSpace(licenseType))
                                            {
                                                int licenseOption = GetOptionSetValue(service, "incident", "new_licensetype", normalizedLicenseType);
                                                incident["new_licensetype"] = new OptionSetValue(licenseOption);
                                            }

                                            // ➕ Set new_compliancetype if sector is Compliance
                                            else if ((department.Equals("Compliance", StringComparison.OrdinalIgnoreCase) ||
                                                      department.Equals("الامتثال", StringComparison.OrdinalIgnoreCase)) &&
                                                     !string.IsNullOrWhiteSpace(complianceType))
                                            {
                                                int complianceOption = GetOptionSetValue(service, "incident", "new_compliancetype", complianceType);
                                                incident["new_compliancetype"] = new OptionSetValue(complianceOption);
                                            }
                                        }


                                        // Optional: Set Reason for Appointment (Text field)
                                        //if (!string.IsNullOrWhiteSpace(reason))
                                        //{
                                        //    incident["new_appointmentreason"] = reason; // Replace with correct schema if needed
                                        //}
                                        service.Create(incident);
                                        tracer.Trace($"✅ Case created successfully.");
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
            catch (Exception ex)
            {
                tracer.Trace("❌ Plugin Exception: " + ex.ToString());
                throw;
            }
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

        static string NormalizeAppointmentType(string rawValue)
        {
            var arabicToEnglish = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "مكالمة فيديو", "Video Call" },
        { "مقابلة مسؤول", "Meeting with Officer" },
        { "زيارة", "Visit" },
        { "أخرى", "Other" }
    };

            rawValue = rawValue.Trim();

            // If Arabic, translate to English
            if (Regex.IsMatch(rawValue, @"\p{IsArabic}"))
            {
                return arabicToEnglish.ContainsKey(rawValue) ? arabicToEnglish[rawValue] : "Other";
            }

            // If English, return as-is or mapped to consistent casing
            return arabicToEnglish
                .Where(kvp => string.Equals(kvp.Value, rawValue, StringComparison.OrdinalIgnoreCase))
                .Select(kvp => kvp.Value)
                .FirstOrDefault() ?? "Other";
        }
        static string NormalizeSector(string rawValue)
        {
            var arabicToEnglish = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "الرخص", "Licensing" },
        { "الامتثال", "Compliance" },
        { "تجربة العميل", "Customer Experience" }
    };

            rawValue = rawValue.Trim();

            // If Arabic, translate to English
            if (Regex.IsMatch(rawValue, @"\p{IsArabic}"))
            {
                return arabicToEnglish.ContainsKey(rawValue) ? arabicToEnglish[rawValue] : "Other";
            }

            // If English, return consistent casing
            return arabicToEnglish
                .Where(kvp => string.Equals(kvp.Value, rawValue, StringComparison.OrdinalIgnoreCase))
                .Select(kvp => kvp.Value)
                .FirstOrDefault() ?? "Other";
        }
        static string NormalizeLicenseType(string rawValue)
        {
            var arabicToEnglish = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "رخص الاستكشاف", "Exploration Licenses" },
        { "رخص البناء", "BMQ Licenses" },
        { "رخص المناجم الصغيرة", "Mining and Small Mine Licenses" }
    };

            rawValue = rawValue.Trim();

            // If Arabic, translate to English
            if (Regex.IsMatch(rawValue, @"\p{IsArabic}"))
            {
                return arabicToEnglish.ContainsKey(rawValue) ? arabicToEnglish[rawValue] : "Other";
            }

            // If English, return consistent casing
            return arabicToEnglish
                .Where(kvp => string.Equals(kvp.Value, rawValue, StringComparison.OrdinalIgnoreCase))
                .Select(kvp => kvp.Value)
                .FirstOrDefault() ?? "Other";
        }

    }
}

