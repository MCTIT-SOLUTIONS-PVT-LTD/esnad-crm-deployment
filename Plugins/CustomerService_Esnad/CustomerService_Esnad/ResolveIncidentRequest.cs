using Microsoft.Xrm.Sdk;

namespace CustomerService_Esnad
{
    internal class ResolveIncidentRequest
    {
        public Entity IncidentResolution { get; internal set; }
        public int Status { get; internal set; }
    }
}