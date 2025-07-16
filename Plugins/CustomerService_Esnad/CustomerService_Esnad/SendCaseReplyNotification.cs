using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomerService_Esnad
{
    public class SendCaseReplyNotification : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracing.Trace("🔔 Plugin execution started.");

            try
            {
                if (!context.InputParameters.Contains("CaseId") || !(context.InputParameters["CaseId"] is EntityReference caseRef))
                    throw new InvalidPluginExecutionException("Missing or invalid 'CaseId' input parameter.");

                if (!context.InputParameters.Contains("TeamId") || !(context.InputParameters["TeamId"] is EntityReference teamRef))
                    throw new InvalidPluginExecutionException("Missing or invalid 'TeamId' input parameter.");

                var caseId = caseRef.Id;
                var teamId = teamRef.Id;

                // Get case title
                var caseEntity = service.Retrieve("incident", caseId, new ColumnSet("title"));
                string caseTitle = caseEntity.GetAttributeValue<string>("title") ?? "Unknown";

                // Get all users in the team
                var teamUsersQuery = new QueryExpression("teammembership")
                {
                    ColumnSet = new ColumnSet("systemuserid"),
                    Criteria = new FilterExpression
                    {
                        Conditions = {
                            new ConditionExpression("teamid", ConditionOperator.Equal, teamId)
                        }
                    }
                };

                var userIds = service.RetrieveMultiple(teamUsersQuery)
                    .Entities.Select(e => e.GetAttributeValue<Guid>("systemuserid")).Distinct().ToList();

                if (!userIds.Any())
                {
                    tracing.Trace("❌ No users found in the team.");
                    return;
                }

                // Build 'To' recipients
                var toParties = userIds.Select(uid => new Entity("activityparty")
                {
                    ["partyid"] = new EntityReference("systemuser", uid)
                }).ToList();

                // Get CRM Admin user
                var crmAdmin = service.RetrieveMultiple(new QueryExpression("systemuser")
                {
                    ColumnSet = new ColumnSet("systemuserid", "internalemailaddress"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("domainname", ConditionOperator.Equal, "CRM-ESNAD\\crmadmin"),
                            new ConditionExpression("accessmode", ConditionOperator.Equal, 0)
                        }
                    }
                }).Entities.FirstOrDefault();

                if (crmAdmin == null || !crmAdmin.Contains("internalemailaddress"))
                    throw new InvalidPluginExecutionException("CRM Admin user not found or missing email.");

                var fromParty = new Entity("activityparty")
                {
                    ["partyid"] = new EntityReference("systemuser", crmAdmin.Id)
                };

                // Build email
                string orgUrl = GetOrgURL(service, tracing);
                string caseUrl = $"{orgUrl}{caseId}";
                string caseTitleHtml = $"<a href='{caseUrl}' style='color:#0078d4; font-weight:bold;'>{caseTitle}</a>";

                string emailBody = $@"
<html>
  <body>
    <p><img src='http://d365.crm-esnad.com/' alt='CRM Logo' style='max-width: 200px;' /></p>
    <p>📝 <strong>Customer has responded to the ticket:</strong> {caseTitleHtml}</p>
    <p>يرجى مراجعة الرد واتخاذ الإجراءات اللازمة.</p>
  </body>
</html>";

                var email = new Entity("email")
                {
                    ["subject"] = $"Customer Response for - {caseTitle}",
                    ["description"] = emailBody,
                    ["directioncode"] = true,
                    ["from"] = new EntityCollection(new[] { fromParty }),
                    ["to"] = new EntityCollection(toParties),
                    ["regardingobjectid"] = new EntityReference("incident", caseId),
                    ["statuscode"] = new OptionSetValue(1) // Draft
                };

                Guid emailId = service.Create(email);
                tracing.Trace("✅ Email created. ID: " + emailId);

                var sendRequest = new SendEmailRequest
                {
                    EmailId = emailId,
                    IssueSend = true,
                    TrackingToken = ""
                };

                service.Execute(sendRequest);
                tracing.Trace("✅ Email sent via SendEmailRequest.");

                // Update case
                var updateCase = new Entity("incident", caseId)
                {
                    ["new_copycaseguid"] = caseId.ToString()
                };
                service.Update(updateCase);
                tracing.Trace("✅ Case updated with new_copycaseguid.");

            }
            catch (Exception ex)
            {
                tracing.Trace("❌ Exception: " + ex.ToString());
                throw new InvalidPluginExecutionException("Error in SendCaseReplyNotificationPlugin.", ex);
            }

            tracing.Trace("🏁 Plugin execution completed.");
        }

        private string GetOrgURL(IOrganizationService service, ITracingService tracing)
        {
            var query = new QueryExpression("new_environmentvariable")
            {
                ColumnSet = new ColumnSet("new_value"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("new_name", ConditionOperator.Equal, "OrgURL")
                    }
                }
            };

            var result = service.RetrieveMultiple(query);
            if (result.Entities.Count > 0)
            {
                return result.Entities[0].GetAttributeValue<string>("new_value");
            }

            tracing.Trace("❌ OrgURL environment variable not found.");
            throw new InvalidPluginExecutionException("OrgURL environment variable missing.");
        }
    }
}