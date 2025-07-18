using System;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;

namespace CustomerService_Esnad
{
    public class ResolveCasePlugin : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            // Setup services
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            ITracingService tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = factory.CreateOrganizationService(context.UserId);

            tracing.Trace("📌 ResolveCasePlugin started.");

            try
            {
                // Validate input
                if (!context.InputParameters.Contains("CaseId") || !(context.InputParameters["CaseId"] is EntityReference caseRef))
                {
                    tracing.Trace("⚠️ CaseId is missing or invalid.");
                    return;
                }

                if (caseRef.LogicalName != "incident")
                {
                    tracing.Trace("⚠️ CaseId is not an incident entity.");
                    return;
                }

                Guid caseId = caseRef.Id;
                tracing.Trace($"✅ Resolving Case: {caseId}");

                // Create incident resolution after the sending Email
                Entity resolution = new Entity("incidentresolution");
                resolution["subject"] = "Auto-closed via workflow action";
                resolution["incidentid"] = new EntityReference("incident", caseId);

                // Execute CloseIncidentRequest
                var closeRequest = new CloseIncidentRequest
                {
                    IncidentResolution = resolution,
                    Status = new OptionSetValue(5) // 5 = Problem Solved (valid status reason for Resolved)
                };

                service.Execute(closeRequest);

                // Now update the incident status to "Resolved"
                Entity caseToUpdate = new Entity("incident", caseId);
                caseToUpdate["statuscode"] = new OptionSetValue(5); // 5 represents "Resolved" status
                service.Update(caseToUpdate);
                tracing.Trace("✅ Case successfully resolved.");
            }
            catch (Exception ex)
            {
                tracing.Trace("❌ Exception: " + ex.ToString());
                throw new InvalidPluginExecutionException("Failed to resolve case via plugin.", ex);
            }
        }


    }
}
