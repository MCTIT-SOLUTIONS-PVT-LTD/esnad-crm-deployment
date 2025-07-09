using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CustomerService_Esnad
{


    public class CreateOnBPF : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.OutputParameters.Contains("id") && context.OutputParameters["id"] is Guid caseId)
            {
                // Step 1: Get the active BPF for Case
                var bpfQuery = new QueryExpression("workflow")
                {
                    ColumnSet = new ColumnSet("workflowid", "name"),
                    Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("category", ConditionOperator.Equal, 4), // BPF
                        new ConditionExpression("primaryentity", ConditionOperator.Equal, "incident"),
                        new ConditionExpression("statecode", ConditionOperator.Equal, 1) // Active
                    }
                }
                };

                var bpf = service.RetrieveMultiple(bpfQuery).Entities.FirstOrDefault();
                if (bpf == null) return;

                // Step 2: Get BPF stages (without ordering)
                var stageQuery = new QueryExpression("processstage")
                {
                    ColumnSet = new ColumnSet("processstageid", "stagename", "processid"),
                    Criteria =
                {
                    Conditions =
                    {
                        new ConditionExpression("processid", ConditionOperator.Equal, bpf.Id)
                    }
                }
                };

                var stages = service.RetrieveMultiple(stageQuery).Entities;
                if (stages.Count < 2) return;

                // Fallback: Just pick the second stage in the list
                var secondStage = stages[1];

                // Step 3: Create BPF instance
                var bpfInstance = new Entity("ticketprocessflow");
                bpfInstance["bpf_incidentid"] = new EntityReference("incident", caseId);
                bpfInstance["processid"] = bpf.Id;
                bpfInstance["activestageid"] = secondStage.ToEntityReference();
                bpfInstance["traversedpath"] = secondStage.Id.ToString();
                Guid bpfInstanceId = service.Create(bpfInstance);

                // Step 4: Update Case
                var caseUpdate = new Entity("incident", caseId)
                {
                    ["activestageid"] = secondStage.ToEntityReference(),
                    ["processid"] = bpf.Id,
                    ["traversedpath"] = secondStage.Id.ToString()
                };
                service.Update(caseUpdate);
            }
        }
    }
}

