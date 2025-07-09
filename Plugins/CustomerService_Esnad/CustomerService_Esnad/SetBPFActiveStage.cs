using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;

namespace CustomerService_Esnad
{


    public class SetBPFActiveStage : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                // Ensure the plugin is running on the "Update" message for the correct entity (case in this case)
                if (context.MessageName != "Update")
                {
                    return;
                }

                Entity entity = (Entity)context.InputParameters["Target"];

                // Check if the necessary field (new_stagesbpf) is present
                if (!entity.Attributes.Contains("new_stagesbpf"))
                {
                    throw new InvalidPluginExecutionException("The field new_stagesbpf is missing from the case entity.");
                }

                OptionSetValue selectedValue = (OptionSetValue)entity["new_stagesbpf"];
                string targetStageName = GetStageNameByValue(selectedValue.Value);

                if (string.IsNullOrEmpty(targetStageName))
                {
                    throw new InvalidPluginExecutionException("Invalid stage value selected.");
                }

                string targetStageId = GetStageGuidByName(service, targetStageName);

                if (string.IsNullOrEmpty(targetStageId))
                {
                    throw new InvalidPluginExecutionException($"Could not find a stage with the name '{targetStageName}'.");
                }

                // Retrieve the BPF record
                Entity bpfRecord = GetBPFRecord(service, entity.Id);
                if (bpfRecord == null)
                {
                    throw new InvalidPluginExecutionException("No BPF record found for the provided case.");
                }

                // Set the active stage for the BPF record
                bpfRecord["activestageid"] = new EntityReference("processstage", new Guid(targetStageId));

                // Update the BPF record
                service.Update(bpfRecord);
                Console.WriteLine($"✅ Successfully moved to stage '{targetStageName}' with GUID: {targetStageId}");
            }
            catch (InvalidPluginExecutionException ex)
            {
                // Handle expected errors with a specific message
                throw new InvalidPluginExecutionException($"Plugin Error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // Catch unexpected errors and log them
                string errorMessage = $"Unexpected error occurred in the SetBPFActiveStage plugin: {ex.Message}";
                throw new InvalidPluginExecutionException(errorMessage, ex);
            }
        }

        private string GetStageNameByValue(int value)
        {
            switch (value)
            {
                case 1:
                    return "Solution Verification";
                case 2:
                    return "Return to Customer";
                case 3:
                    return "Processing- Department";
                default:
                    return null;
            }
        }

        private string GetStageGuidByName(IOrganizationService service, string stageName)
        {
            // Query to find the process stage by name
            QueryExpression query = new QueryExpression("processstage")
            {
                ColumnSet = new ColumnSet("processstageid", "name")
            };

            query.Criteria.AddCondition("name", ConditionOperator.Equal, stageName);

            EntityCollection result = service.RetrieveMultiple(query);

            if (result.Entities.Count > 0)
            {
                return result.Entities[0].GetAttributeValue<Guid>("processstageid").ToString();
            }

            return null;
        }

        private Entity GetBPFRecord(IOrganizationService service, Guid caseId)
        {
            // Query to get the BPF record associated with the case
            QueryExpression query = new QueryExpression("phonetocaseprocess")
            {
                ColumnSet = new ColumnSet("phonetocaseprocessid", "activestageid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                {
                    new ConditionExpression("_incidentid_value", ConditionOperator.Equal, caseId)
                }
                }
            };

            EntityCollection result = service.RetrieveMultiple(query);

            if (result.Entities.Count > 0)
            {
                return result.Entities[0];
            }

            return null;
        }
    }
}