using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CustomerService_Esnad
{
    public class SLALevel3 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Get the context
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            tracing.Trace("SLALevel1Escalation Plugin execution started.");

            try
            {
                // Get Case ID from InputParameters
                if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is EntityReference caseRef))
                {
                    tracing.Trace("CaseId not found in input parameters.");
                    return;
                }

                Guid caseId = caseRef.Id;
                tracing.Trace($"Processing Case ID: {caseId}");

                // Retrieve Case details
                Entity caseEntity = service.Retrieve("incident", caseId, new ColumnSet("ownerid", "title"));
                if (!caseEntity.Contains("ownerid"))
                {
                    tracing.Trace("Case does not have an owner. Exiting.");
                    return;
                }

                string caseTitle = caseEntity.GetAttributeValue<string>("title") ?? "(No Title)";
                EntityReference ownerRef = caseEntity.GetAttributeValue<EntityReference>("ownerid");
                tracing.Trace($"Case Owner: {ownerRef.Name}, Type: {ownerRef.LogicalName}");

                // Fetch crmadmin as sender
                Entity crmAdminUser = GetCRMAdminUser(service);
                if (crmAdminUser == null)
                    throw new InvalidPluginExecutionException("CRM Admin user not found or missing email.");

                var fromParty = new Entity("activityparty")
                {
                    ["partyid"] = new EntityReference("systemuser", crmAdminUser.Id)
                };

                string orgURL = GetOrgURL(service);
                string caseUrl = $"{orgURL}{caseId}";

                if (ownerRef.LogicalName == "team")
                {
                    tracing.Trace("Owner is a Team. Sending email to Sector Head in this team.");
                    SendEmailToTeam(service, crmAdminUser, fromParty, caseId, caseTitle, ownerRef, ownerRef.Id, caseUrl, tracing, ownerRef.Name);
                }
                else if (ownerRef.LogicalName == "systemuser")
                {
                    tracing.Trace("Owner is a User. Fetching user's teams...");
                    var teams = GetUserTeams(service, ownerRef.Id, tracing);
                    tracing.Trace($"Found {teams.Count} teams for user.");

                    foreach (var team in teams)
                    {
                        tracing.Trace($"Processing team: {team.GetAttributeValue<string>("name")}");
                        SendEmailToTeam(service, crmAdminUser, fromParty, caseId, caseTitle, ownerRef, team.Id, caseUrl, tracing, team.GetAttributeValue<string>("name"));
                    }
                }

                tracing.Trace("SLALevel1Escalation Plugin execution completed.");
            }
            catch (Exception ex)
            {
                tracing.Trace("Error: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed in SLALevel1Escalation plugin.", ex);
            }
        }

        private void SendEmailToTeam(IOrganizationService service, Entity crmAdminUser, Entity fromParty, Guid caseId, string caseTitle, EntityReference ownerRef, Guid teamId, string caseUrl, ITracingService tracing, string teamName)
        {
            var users = GetSectorHeadInTeam(service, teamId, tracing);
            if (users.Count == 0)
            {
                tracing.Trace($"No Department Manager found in team: {teamId}");
                return;
            }

            var toParties = users.Select(u => new Entity("activityparty")
            {
                ["partyid"] = new EntityReference("systemuser", u.Id)
            }).ToList();

            tracing.Trace($"Creating email for team: {teamName}");

            string subject = $"[SLA Escalation Level 3] {teamName} - Case Breach Alert";
            string imageUrl = "http://d365.crm-esnad.com/";

            var email = new Entity("email")
            {
                ["subject"] = subject,
                ["description"] = $@"
        <html>
        <body>
            <p><img src='{imageUrl}' alt='CRM Logo' style='width:200px; margin-bottom:10px;' /></p>
            <p>Dear Sector Head Team,<br/><br/></p>
            <p>This is to inform you that the following case has breached its SLA threshold:</p>
            <p>Please review: <a href='{caseUrl}' style='color:#0078d4; font-weight:bold;'>{caseTitle}</a></p>
            <p><strong>Assigned Agent:</strong> {ownerRef.Name}</p>
            <br/>
            <p>Thank you,</p>
            <p>Best regards,</p>
            <p>Support Escalation Team</p>
        </body>
        </html>",
                ["directioncode"] = true,
                ["from"] = new EntityCollection(new[] { fromParty }),
                ["to"] = new EntityCollection(toParties),
                ["regardingobjectid"] = new EntityReference("incident",caseId),
                ["statuscode"] = new OptionSetValue(1) // Draft
            };


            Guid emailId = service.Create(email);
            tracing.Trace($"Email created for team {teamName}. ID: {emailId}");

            var sendRequest = new OrganizationRequest("SendEmail");
            sendRequest["EmailId"] = emailId;
            sendRequest["IssueSend"] = true;
            sendRequest["TrackingToken"] = "";

            service.Execute(sendRequest);
            tracing.Trace($"Email sent to team {teamName} successfully.");
        }

        private List<Entity> GetUserTeams(IOrganizationService service, Guid userId, ITracingService tracing)
        {
            var fetchXml = $@"
            <fetch>
              <entity name='team'>
                <attribute name='name'/>
                <attribute name='teamid'/>
                <link-entity name='teammembership' from='teamid' to='teamid' intersect='true'>
                  <filter>
                    <condition attribute='systemuserid' operator='eq' value='{userId}'/>
                  </filter>
                </link-entity>
              </entity>
            </fetch>";

            var result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            tracing.Trace($"Found {result.Entities.Count} teams for user {userId}.");
            return result.Entities.ToList();
        }

        private List<Entity> GetSectorHeadInTeam(IOrganizationService service, Guid teamId, ITracingService tracing)
        {
            var fetchXml = $@"
            <fetch>
              <entity name='systemuser'>
                <attribute name='systemuserid'/>
                <attribute name='internalemailaddress'/>
                <filter>
                  <condition attribute='accessmode' operator='eq' value='0' />
                </filter>
                <link-entity name='teammembership' from='systemuserid' to='systemuserid' link-type='inner'>
                  <filter>
                    <condition attribute='teamid' operator='eq' value='{teamId}' />
                  </filter>
                </link-entity>
                <link-entity name='position' from='positionid' to='positionid' link-type='inner'>
                  <filter>
                    <condition attribute='name' operator='eq' value='Sector Head' />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";

            var result = service.RetrieveMultiple(new FetchExpression(fetchXml));
            tracing.Trace($"Found {result.Entities.Count} Sector Head  in team {teamId}.");
            return result.Entities.ToList();
        }

        private Entity GetCRMAdminUser(IOrganizationService service)
        {
            var query = new QueryExpression("systemuser")
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
            };

            return service.RetrieveMultiple(query).Entities.FirstOrDefault();
        }

        private string GetOrgURL(IOrganizationService service)
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

            EntityCollection result = service.RetrieveMultiple(query);
            if (result.Entities.Count > 0)
                return result.Entities[0].GetAttributeValue<string>("new_value");

            throw new InvalidPluginExecutionException("OrgURL environment variable not found.");
        }
    }
}
