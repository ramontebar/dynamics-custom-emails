using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.Messages
{

    [System.Runtime.Serialization.DataContractAttribute(Namespace = "http://schemas.microsoft.com/xrm/2011/dxtools/")]
    [Microsoft.Xrm.Sdk.Client.ResponseProxyAttribute("dxtools_SendCustomEmail")]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("CrmSvcUtil", "6.0.0001.0061")]
    public partial class dxtools_SendCustomEmailResponse : Microsoft.Xrm.Sdk.OrganizationResponse
    {

        public Microsoft.Xrm.Sdk.EntityReference NewCreatedEmail
        {
            get
            {
                if (this.Results.Contains("NewCreatedEmail"))
                {
                    return ((Microsoft.Xrm.Sdk.EntityReference)(this.Results["NewCreatedEmail"]));
                }
                else
                {
                    return default(Microsoft.Xrm.Sdk.EntityReference);
                }
            }
            set
            {
                this.Results["NewCreatedEmail"] = value;
            }
        }

        public String SentEmailSubject
        {
            get
            {
                if (this.Results.Contains("SentEmailSubject"))
                {
                    return this.Results["SentEmailSubject"] as string;
                }
                else
                {
                    return default(string);
                }
            }
            set
            {
                this.Results["SentEmailSubject"] = value;
            }
        }

        public dxtools_SendCustomEmailResponse()
        {
        }

    }
}

