using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace CustomerService_Esnad
    {
        public class MoveBpfToApprovalStage : IPlugin
        {
            public void Execute(IServiceProvider serviceProvider)
            {
                // Services
                var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
                var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
                var service = factory.CreateOrganizationService(context.UserId);
                var tracing = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

                tracing.Trace("🔁 MoveCaseToApprovalStagePlugin triggered.");

                try
                {
                    // Validate input
                    if (!context.InputParameters.Contains("Target") || !(context.InputParameters["Target"] is Entity target) || target.LogicalName != "incident")
                    {
                        tracing.Trace("❌ Invalid Target Entity. Exiting.");
                        return;
                    }

                    Guid caseId = target.Id;
                    tracing.Trace($"📌 Case ID: {caseId}");

                    // Step 1: Get BPF for the Case
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
                    {
                        tracing.Trace("❌ No BPF found for the given case.");
                        return;
                    }

                    var bpf = bpfResult.Entities.First();
                    var bpfId = bpf.Id;
                    var bpfName = bpf.GetAttributeValue<string>("name");
                    tracing.Trace($"✅ Found BPF: {bpfName}, ID: {bpfId}");

                    // Step 2: Set known stage ID for 'Approval And Forwarding'
                    var approvalStageId = new Guid("92a6721b-d465-4d36-aef7-e8822d7a5a6a");
                    tracing.Trace($"📌 Using static Stage ID: {approvalStageId}");

                    // Step 3: Update the BPF stage
                    var updateBpf = new Entity("phonetocaseprocess", bpfId)
                    {
                        ["activestageid"] = new EntityReference("processstage", approvalStageId)
                    };

                    service.Update(updateBpf);
                    tracing.Trace("✅ BPF successfully updated to 'Approval And Forwarding' stage.");

                // ✅ Step 3: Save the Case record to trigger re-evaluation
                var updateCase = new Entity("incident", caseId);
                //updateCase["description"] = target.GetAttributeValue<string>("description"); // optional: reassign same data to simulate change
                service.Update(updateCase);
                tracing.Trace("✅ Case record updated/saved.");
            }
                catch (Exception ex)
                {
                    tracing.Trace($"❌ Exception: {ex}");
                    throw new InvalidPluginExecutionException("Failed to move BPF to 'Approval And Forwarding' stage.", ex);
                }

                tracing.Trace("🏁 MoveCaseToApprovalStagePlugin completed.");
            }
        }
    }

