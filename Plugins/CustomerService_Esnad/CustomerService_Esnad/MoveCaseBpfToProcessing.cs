using System;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace CustomerService_Esnad
{
    public class MoveCaseBpfToProcessing : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracing.Trace("🔁 MoveCaseBpfToProcessingPlugin started.");

            try
            {
                if (!context.InputParameters.Contains("CaseId") || !(context.InputParameters["CaseId"] is EntityReference caseRef))
                    throw new InvalidPluginExecutionException("❌ Input parameter 'CaseId' is missing or invalid.");

                var caseId = caseRef.Id;
                tracing.Trace($"📌 CaseId received: {caseId}");

                // Step 1: Retrieve BPF for the Case
                tracing.Trace("🔍 Retrieving PhoneToCaseProcess linked to the Case...");
                var bpfQuery = new QueryExpression("phonetocaseprocess")
                {
                    ColumnSet = new ColumnSet("businessprocessflowinstanceid", "processid", "activestageid", "name"),
                    Criteria =
                    {
                        Conditions =
                        {
                            new ConditionExpression("incidentid", ConditionOperator.Equal, caseId)
                        }
                    }
                };

                var bpfResult = service.RetrieveMultiple(bpfQuery);
                if (!bpfResult.Entities.Any())
                    throw new InvalidPluginExecutionException("❌ No PhoneToCaseProcess record found for the given Case.");

                var bpf = bpfResult.Entities.First();
                var bpfId = bpf.Id;
                var bpfName = bpf.GetAttributeValue<string>("name");
                tracing.Trace($"✅ BPF found - ID: {bpfId}, Name: {bpfName}");

                // Step 2: Set hardcoded 'Processing' stage ID
                var processingStageId = new Guid("91153307-982f-479d-af7f-73048b80e52c");
                tracing.Trace($"📌 Using hardcoded 'Processing' stage ID: {processingStageId}");

                // Step 3: Update the BPF to move to the specified stage
                var updateBpf = new Entity("phonetocaseprocess", bpfId)
                {
                    ["activestageid"] = new EntityReference("processstage", processingStageId)
                };

                service.Update(updateBpf);
                tracing.Trace("✅ BPF stage updated to 'Processing'.");

                // Step 4: Save the Case record to trigger reevaluation
                var updateCase = new Entity("incident", caseId)
                {
                    ["statuscode"] = new OptionSetValue(100000006)
                };
                service.Update(updateCase);
                tracing.Trace("✅ Case record saved to trigger business logic.");

            }
            catch (Exception ex)
            {
                tracing.Trace($"❌ Exception occurred: {ex}");
                throw new InvalidPluginExecutionException("An error occurred in MoveCaseBpfToProcessingPlugin.", ex);
            }

            tracing.Trace("🏁 MoveCaseBpfToProcessingPlugin completed.");
        }
    }
}
