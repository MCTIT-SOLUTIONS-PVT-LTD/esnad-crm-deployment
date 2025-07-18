﻿using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomerService_Esnad
{


    public class SLALeval1 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);

            tracing.Trace("NotifySpecializedAdminsPlugin execution started.");

            try
            {
                // Input validation
                if (!context.InputParameters.Contains("CaseId") || !(context.InputParameters["CaseId"] is EntityReference caseRef))
                {
                    tracing.Trace("CaseId parameter missing or invalid.");
                    return;
                }

                // Retrieve Case
                Entity caseEntity;
                try
                {
                    caseEntity = service.Retrieve("incident", caseRef.Id, new ColumnSet("ownerid", "title"));
                }
                catch (Exception ex)
                {
                    tracing.Trace("Failed to retrieve case: " + ex.Message);
                    throw new InvalidPluginExecutionException("Error retrieving case record.", ex);
                }

                if (!caseEntity.Attributes.Contains("ownerid"))
                {
                    tracing.Trace("Owner not found on case.");
                    return;
                }

                string caseTitle = caseEntity.GetAttributeValue<string>("title") ?? "(No Title)";
                var ownerRef = caseEntity.GetAttributeValue<EntityReference>("ownerid");
                var userIds = new HashSet<Guid>();

                if (ownerRef.LogicalName == "team")
                {
                    var users = GetSpecializedAdminsInTeam(service, ownerRef.Id, tracing);
                    foreach (var u in users) userIds.Add(u.Id);
                }
                else if (ownerRef.LogicalName == "systemuser")
                {
                    var teamIds = GetUserTeams(service, ownerRef.Id, tracing);
                    foreach (var teamId in teamIds)
                    {
                        var users = GetSpecializedAdminsInTeam(service, teamId, tracing);
                        foreach (var u in users) userIds.Add(u.Id);
                    }
                }

                if (!userIds.Any())
                {
                    tracing.Trace("No CRM Officer users found.");
                    return;
                }

                var toParties = userIds.Select(id => new Entity("activityparty")
                {
                    ["partyid"] = new EntityReference("systemuser", id)
                }).ToList();

                // Fetch sender user: "CRM-ESNAD\\crmadmin"
                Entity crmAdminUser = service.RetrieveMultiple(new QueryExpression("systemuser")
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

                if (crmAdminUser == null)
                    throw new InvalidPluginExecutionException("crmadmin user not found or inactive.");

                if (!crmAdminUser.Contains("internalemailaddress"))
                    throw new InvalidPluginExecutionException("crmadmin user does not have a valid email.");

                var fromParty = new Entity("activityparty")
                {
                    ["partyid"] = new EntityReference("systemuser", crmAdminUser.Id)
                };
                string imageUrl = "http://d365.crm-esnad.com/"; // Use HTTPS if possible
                //string caseUrl = $"https://d365.crm-esnad.com/main.aspx?appid=0d3f8ee3-bd6f-4d2a-8205-8b8d5021b809&pagetype=entityrecord&etn=incident&id={caseRef.Id}";
                // Get OrgURL from the environment variable entity
                string orgURL = GetOrgURL(service);
                // Create the final case URL by concatenating OrgURL and the Case Id
                string caseUrl = $"{orgURL}{caseRef.Id}";  // Concatenate the OrgURL and Case Id
                // Create Email
                var email = new Entity("email")
                {
                    ["subject"] = "Level 1 Escalation: Immediate Attention Needed for Case",
                    ["description"] = $@"
        <html>
        <body>
            <p><img src='{imageUrl}' alt='CRM Logo' style='width:200px; margin-bottom:10px;' /></p>
            <p>Dear Core Support Team,<br/><br/></p>
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
                    ["regardingobjectid"] = new EntityReference("incident", caseRef.Id),
                    ["statuscode"] = new OptionSetValue(1) // Draft
                };
                Guid emailId = service.Create(email);
                tracing.Trace("Email created. ID: " + emailId);

                // Force send the email
                var sendRequest = new OrganizationRequest("SendEmail");
                sendRequest["EmailId"] = emailId;
                sendRequest["IssueSend"] = true;
                sendRequest["TrackingToken"] = "";

                service.Execute(sendRequest);
                tracing.Trace("Email sent via SendEmailRequest.");
            }
            catch (Exception ex)
            {
                tracing.Trace("NotifySpecializedAdminsPlugin error: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed to notify CRM Officer.", ex);
            }
        }

        private List<Guid> GetUserTeams(IOrganizationService service, Guid userId, ITracingService tracing)
        {
            try
            {
                var query = new QueryExpression("teammembership")
                {
                    ColumnSet = new ColumnSet("teamid"),
                    Criteria = new FilterExpression
                    {
                        Conditions = { new ConditionExpression("systemuserid", ConditionOperator.Equal, userId) }
                    }
                };

                return service.RetrieveMultiple(query).Entities
                    .Select(e => e.GetAttributeValue<Guid>("teamid"))
                    .ToList();
            }
            catch (Exception ex)
            {
                tracing.Trace("Error in GetUserTeams: " + ex.Message);
                throw;
            }
        }

        private List<Entity> GetSpecializedAdminsInTeam(IOrganizationService service, Guid teamId, ITracingService tracing)
        {
            try
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
                    <condition attribute='name' operator='eq' value='Specialized Dept. Officer' />
                  </filter>
                </link-entity>
              </entity>
            </fetch>";

                var result = service.RetrieveMultiple(new FetchExpression(fetchXml));
                tracing.Trace($"Found {result.Entities.Count} CRM Officer users in team {teamId}");
                return result.Entities.ToList();
            }
            catch (Exception ex)
            {
                tracing.Trace("Error in GetSpecializedAdminsInTeam: " + ex.Message);
                throw;
            }
        }
        private string GetOrgURL(IOrganizationService service)
        {
            // Create a query to find the record where "new_name" equals "OrgURL"
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

            // Retrieve the record
            EntityCollection result = service.RetrieveMultiple(query);

            // Check if the result contains any matching records
            if (result.Entities.Count > 0)
            {
                // Get the "new_value" field value from the first matching record
                string orgURL = result.Entities[0].GetAttributeValue<string>("new_value");
                return orgURL;
            }
            else
            {
                throw new InvalidPluginExecutionException("No record found for 'OrgURL' in 'new_environmentvariable' entity.");
            }
        }
    }
}
