using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Web.Http;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using TicketSystemApi.Services;

[RoutePrefix("api/cases")]
public class CaseReportController : ApiController
{
    [HttpGet]
    [Route("report")]
    public IHttpActionResult GetCases(string filter = "all", int page = 1, int? pageSize = null)
    {
        try
        {
            var allowedTokens = new List<string>
            {
                System.Configuration.ConfigurationManager.AppSettings["ReportDataToken"],
                System.Configuration.ConfigurationManager.AppSettings["ReportDataToken1"]
            };

            string token = "";

            if (Request.Headers.Authorization != null && Request.Headers.Authorization.Scheme == "Bearer")
            {
                token = Request.Headers.Authorization.Parameter;
            }
            else if (Request.GetQueryNameValuePairs().Any(kvp => kvp.Key == "token"))
            {
                token = Request.GetQueryNameValuePairs().First(kvp => kvp.Key == "token").Value;
            }

            if (!allowedTokens.Contains(token))
            {
                return Unauthorized();
            }

            ICrmService crm = new CrmService();
            var service = crm.GetService();

            var query = new QueryExpression("incident")
            {
                ColumnSet = new ColumnSet(
                    "ticketnumber", "createdon", "modifiedon", "statuscode", "prioritycode",
                    "new_ticketclosuredate", "new_description", "new_ticketsubmissionchannel",
                    "new_businessunitid", "createdby", "modifiedby", "ownerid", "customerid",
                    "new_tickettype", "new_mainclassification", "new_subclassificationitem",
                    "new_isreopened"
                ),
                PageInfo = new PagingInfo
                {
                    PageNumber = page,
                    Count = (!pageSize.HasValue || pageSize.Value <= 0) ? int.MaxValue : pageSize.Value,
                    PagingCookie = null
                }

            };

            query.Orders.Add(new OrderExpression("createdon", OrderType.Descending));

            var ksaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time");
            var ksaNow = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, ksaTimeZone);
            DateTime ksaStart;

            if (filter.ToLower() == "daily")
            {
                ksaStart = ksaNow.Date;
                query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, ksaStart);
            }
            else if (filter.ToLower() == "weekly")
            {
                ksaStart = ksaNow.Date.AddDays(-7); // ✅ Last 7 days, not week start
                query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, ksaStart);
            }
            else if (filter.ToLower() == "monthly")
            {
                ksaStart = new DateTime(ksaNow.Year, ksaNow.Month, 1);
                query.Criteria.AddCondition("createdon", ConditionOperator.OnOrAfter, ksaStart);
            }

            var result = service.RetrieveMultiple(query);

            var records = result.Entities.Select(e =>
            {
                var slaDetails = GetSlaDetails(service, e.Id);
                var currentStage = MapStatusCodeToStage(e);
                var escalationLevel = GetEscalationLevel(slaDetails);
                var csat = GetCustomerSatisfactionFeedback(service, e.Id);

                return new
                {
                    TicketID = e.GetAttributeValue<string>("ticketnumber"),
                    CreatedBy = e.GetAttributeValue<EntityReference>("createdby")?.Name,
                    AgentName = e.GetAttributeValue<EntityReference>("ownerid")?.Name,
                    CustomerID = e.GetAttributeValue<EntityReference>("customerid")?.Id,
                    CustomerName = e.GetAttributeValue<EntityReference>("customerid")?.Name,
                    CreatedOn = ConvertToKsaTime(e.GetAttributeValue<DateTime?>("createdon")),
                    TicketType = e.GetAttributeValue<EntityReference>("new_tickettype")?.Name,
                    Category = e.GetAttributeValue<EntityReference>("new_tickettype")?.Name,
                    SubCategory1 = e.GetAttributeValue<EntityReference>("new_mainclassification")?.Name,
                    SubCategory2 = e.GetAttributeValue<EntityReference>("new_subclassificationitem")?.Name,
                    Status = e.FormattedValues.Contains("statuscode") ? e.FormattedValues["statuscode"] : null,
                    TicketStatusDateTime = ConvertToKsaTime(e.GetAttributeValue<DateTime?>("modifiedon")),
                    Department = e.Attributes.Contains("new_businessunitid") ? ((EntityReference)e["new_businessunitid"]).Name : null,
                    TicketChannel = e.FormattedValues.Contains("new_ticketsubmissionchannel") ? e.FormattedValues["new_ticketsubmissionchannel"] : null,
                    TotalResolutionTime = (e.Contains("new_ticketclosuredate") && e.Contains("createdon"))
                        ? (e.GetAttributeValue<DateTime>("new_ticketclosuredate") - e.GetAttributeValue<DateTime>("createdon")).ToString(@"hh\:mm\:ss")
                        : null,
                    TotalClosedTime = (e.Contains("modifiedon") && e.Contains("createdon"))
                        ? (e.GetAttributeValue<DateTime>("modifiedon") - e.GetAttributeValue<DateTime>("createdon")).ToString(@"hh\:mm\:ss")
                        : null,
                    Description = e.GetAttributeValue<string>("new_description"),
                    ModifiedBy = e.GetAttributeValue<EntityReference>("modifiedby")?.Name,
                    Priority = e.FormattedValues.Contains("prioritycode") ? e.FormattedValues["prioritycode"] : null,
                    ClosedOn = ConvertToKsaTime(e.GetAttributeValue<DateTime?>("new_ticketclosuredate")),
                    ResolutionDateTime = GetResolutionDateTime(e),
                    CustomerSatisfactionScore = csat.Score,
                    CustomerComment = csat.Comment,
                    AppropriateTimeTaken = csat.AppropriateTimeTaken,
                    IsReopened = string.IsNullOrWhiteSpace(e.GetAttributeValue<string>("new_isreopened")) ? "No" : e.GetAttributeValue<string>("new_isreopened"),
                    SlaViolation = GetSlaViolationStatus(service, e.Id),
                    AssignmentTimeByKPI = slaDetails["AssignmentTimeByKPI"]?.ToString(),
                    ProcessingTimeByKPI = slaDetails["ProcessingTimeByKPI"]?.ToString(),
                    SolutionVerificationTimeByKPI = slaDetails["SolutionVerificationTimeByKPI"]?.ToString(),
                    CurrentStage = currentStage,
                    EscalationLevel = escalationLevel
                };
            }).ToList();

            return Ok(new { Page = page, PageSize = pageSize, Count = records.Count, Records = records });
        }
        catch (Exception ex)
        {
            return InternalServerError(ex);
        }
    }

    private string ConvertToKsaTime(DateTime? utcDateTime)
    {
        if (utcDateTime == null) return null;
        var ksaTimeZone = TimeZoneInfo.FindSystemTimeZoneById("Arab Standard Time");
        var ksaTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(utcDateTime.Value, DateTimeKind.Utc), ksaTimeZone);
        return ksaTime.ToString("yyyy-MM-dd HH:mm:ss");
    }

    private string GetResolutionDateTime(Entity incident)
    {
        var resolvedStatusCodes = new HashSet<int> { 5, 6, 100000003, 100000007, 2000 };
        if (incident.Contains("new_ticketclosuredate") && incident["new_ticketclosuredate"] is DateTime closure)
            return ConvertToKsaTime(closure);

        if (incident.Contains("statuscode") && incident["statuscode"] is OptionSetValue status &&
            resolvedStatusCodes.Contains(status.Value) &&
            incident.Contains("modifiedon") && incident["modifiedon"] is DateTime modified)
            return ConvertToKsaTime(modified);

        return null;
    }

    private Dictionary<string, object> GetSlaDetails(IOrganizationService service, Guid caseId)
    {
        var query = new QueryExpression("slakpiinstance")
        {
            ColumnSet = new ColumnSet("name", "status"),
            Criteria = new FilterExpression
            {
                Conditions = { new ConditionExpression("regarding", ConditionOperator.Equal, caseId) }
            }
        };

        var result = service.RetrieveMultiple(query);
        var details = new Dictionary<string, object>
        {
            { "AssignmentTimeByKPI", null },
            { "ProcessingTimeByKPI", null },
            { "SolutionVerificationTimeByKPI", null }
        };

        foreach (var kpi in result.Entities)
        {
            var name = kpi.GetAttributeValue<string>("name")?.Replace(" ", "").Replace("byKPI", "ByKPI");
            var status = kpi.FormattedValues.Contains("status") ? kpi.FormattedValues["status"] : null;
            if (!string.IsNullOrEmpty(name) && details.ContainsKey(name))
                details[name] = status;
        }

        return details;
    }

    private string GetSlaViolationStatus(IOrganizationService service, Guid caseId)
    {
        var query = new QueryExpression("slakpiinstance")
        {
            ColumnSet = new ColumnSet("status"),
            Criteria = new FilterExpression
            {
                Conditions = { new ConditionExpression("regarding", ConditionOperator.Equal, caseId) }
            }
        };

        var kpiRecords = service.RetrieveMultiple(query);
        foreach (var kpi in kpiRecords.Entities)
        {
            if (kpi.GetAttributeValue<OptionSetValue>("status")?.Value == 3) return "Yes";
        }
        return "No";
    }

    private (int? Score, string Comment, string AppropriateTimeTaken) GetCustomerSatisfactionFeedback(IOrganizationService service, Guid caseId)
    {
        var query = new QueryExpression("new_customersatisfactionscore")
        {
            ColumnSet = new ColumnSet("new_customersatisfactionrating", "new_customersatisfactionscore", "new_wasthetimetakentoprocesstheticketappropri"),
            Criteria = new FilterExpression
            {
                Conditions = { new ConditionExpression("new_csatcase", ConditionOperator.Equal, caseId) }
            }
        };

        var result = service.RetrieveMultiple(query);
        var record = result.Entities.FirstOrDefault();
        if (record == null)
            return (null, null, null);

        var score = record.GetAttributeValue<OptionSetValue>("new_customersatisfactionrating")?.Value;
        var comment = record.GetAttributeValue<string>("new_customersatisfactionscore");

        string appropriateTimeTaken = null;
        if (record.Attributes.Contains("new_wasthetimetakentoprocesstheticketappropri"))
        {
            var timeTaken = record.GetAttributeValue<bool?>("new_wasthetimetakentoprocesstheticketappropri");
            appropriateTimeTaken = timeTaken.HasValue ? (timeTaken.Value ? "Yes" : "No") : null;
        }

        return (score, comment, appropriateTimeTaken);
    }

    private string MapStatusCodeToStage(Entity ticket)
    {
        var statusCode = ticket.GetAttributeValue<OptionSetValue>("statuscode")?.Value;
        switch (statusCode)
        {
            case 100000000: return "Ticket Creation";
            case 100000006: return "Approval and Forwarding";
            case 100000002: return "Solution Verification";
            case 100000008: return "Processing";
            case 1: return "Processing- Department";
            case 100000001: return "Return to Customer";
            case 100000003: return "Ticket Closure";
            case 100000005: return "Ticket Reopen";
            case 5: return "Problem Solved";
            case 1000: return "Information Provided";
            case 6: return "Cancelled";
            case 2000: return "Merged";
            case 100000007: return "Close";
            default: return "Unknown";
        }
    }

    private string GetEscalationLevel(Dictionary<string, object> slaStatuses)
    {
        bool isAssignmentInProgress = string.Equals(slaStatuses["AssignmentTimeByKPI"]?.ToString(), "In Progress", StringComparison.OrdinalIgnoreCase);
        bool isProcessingInProgress = string.Equals(slaStatuses["ProcessingTimeByKPI"]?.ToString(), "In Progress", StringComparison.OrdinalIgnoreCase);
        bool isVerificationInProgress = string.Equals(slaStatuses["SolutionVerificationTimeByKPI"]?.ToString(), "In Progress", StringComparison.OrdinalIgnoreCase);

        if (isAssignmentInProgress && isProcessingInProgress && isVerificationInProgress)
            return "Level 1";
        if (!isAssignmentInProgress && isProcessingInProgress && isVerificationInProgress)
            return "Level 2";
        if (!isAssignmentInProgress && !isProcessingInProgress && isVerificationInProgress)
            return "Level 3";

        return null;
    }
}
