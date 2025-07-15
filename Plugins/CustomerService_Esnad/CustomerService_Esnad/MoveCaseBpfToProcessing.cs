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
            // Services
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);
            var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            tracing.Trace("🔁 MoveCaseBpfToProcessingPlugin started.");

            try
            {
                // Validate and retrieve input parameter
                if (!context.InputParameters.Contains("CaseId") || !(context.InputParameters["CaseId"] is EntityReference caseRef))
                    throw new InvalidPluginExecutionException("❌ Input parameter 'CaseId' is missing or invalid.");

                var caseId = caseRef.Id;
                tracing.Trace($"📌 CaseId received: {caseId}");

                // Step 1: Retrieve BPF (PhoneToCaseProcess) for the Case
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
                var processId = bpf.GetAttributeValue<EntityReference>("processid")?.Id ?? Guid.Empty;
                var bpfName = bpf.GetAttributeValue<string>("name");

                tracing.Trace($"✅ BPF found - ID: {bpfId}, Name: {bpfName}, Process ID: {processId}");

                // Step 2: Retrieve 'Processing' stage for the BPF process
                tracing.Trace("🔍 Retrieving 'Processing' stage from processstage...");
                var stageQuery = new QueryExpression("processstage")
                {
                    ColumnSet = new ColumnSet("processstageid", "stagename"),
                    Criteria =
                    {
                        Conditions =
                        {
                            new ConditionExpression("processid", ConditionOperator.Equal, processId),
                            new ConditionExpression("stagename", ConditionOperator.Equal, "Processing")
                        }
                    }
                };

                var stageResult = service.RetrieveMultiple(stageQuery);
                if (!stageResult.Entities.Any())
                    throw new InvalidPluginExecutionException("❌ 'Processing' stage not found in the process.");

                var processingStageId = stageResult.Entities.First().Id;
                tracing.Trace($"✅ Found 'Processing' stage ID: {processingStageId}");

                // Step 3: Update the active stage of the BPF
                tracing.Trace("📤 Updating BPF's activestageid to 'Processing'...");
                var updateBpf = new Entity("phonetocaseprocess", bpfId)
                {
                    ["activestageid"] = new EntityReference("processstage", processingStageId)
                };

                service.Update(updateBpf);
                tracing.Trace("✅ Successfully moved the BPF to 'Processing' stage.");
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
