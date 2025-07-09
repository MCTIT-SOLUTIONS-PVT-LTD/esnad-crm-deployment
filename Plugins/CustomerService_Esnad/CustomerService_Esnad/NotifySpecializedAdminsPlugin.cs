using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

 public class NotifySpecializedAdminsPlugin : IPlugin
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
                    tracing.Trace("No specialized admin users found.");
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

            string OrgURL = GetOrgURL(service);
            //string caseUrl = $"https://d365.crm-esnad.com/main.aspx?appid=0d3f8ee3-bd6f-4d2a-8205-8b8d5021b809&pagetype=entityrecord&etn=incident&id={caseRef.Id}";
            string imageUrl = "http://d365.crm-esnad.com/"; // Use HTTPS if possible
            string caseUrl = $"{OrgURL}{caseRef.Id}";  // Concatenate the OrgURL and Case Id
            string caseTitleHtml = $"<a href='{caseUrl}' style='color:#0078d4; font-weight:bold;'>{caseTitle}</a>";

            // ✅ Include the image using <img src="">
            string emailBody = $@"
                    <html>
                      <body>
                        <p><img src='{imageUrl}' alt='CRM Logo' style='max-width: 200px;' /></p>
                        <p>تم انشاء تذكرة جديدة رقم {caseTitleHtml}</p>
                        <p>يرجى اعتماد التذكرة وفقاً لاتفاقية مستوى الخدمة</p>
                        
                      </body>
                    </html>";

            var email = new Entity("email")
            {
                ["subject"] = "Case Assigned to your Team",
                ["description"] = emailBody,
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
                throw new InvalidPluginExecutionException("Failed to notify Specialized Admin Staff.", ex);
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
                tracing.Trace($"Found {result.Entities.Count} specialized admin users in team {teamId}");
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


