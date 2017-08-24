using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.Helpers
{
    public class Systemuser
    {
        public static Guid GetCallingUserID(IOrganizationService service)
        {
            WhoAmIRequest whoAmIRequest = new WhoAmIRequest();
            WhoAmIResponse response = service.Execute(whoAmIRequest) as WhoAmIResponse;

            if(response!=null)
                return response.UserId;
            return Guid.Empty;
        }
    }
}
