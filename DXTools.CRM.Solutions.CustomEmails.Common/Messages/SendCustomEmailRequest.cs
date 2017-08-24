using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.Messages
{
    [System.Runtime.Serialization.DataContractAttribute(Namespace = "http://schemas.microsoft.com/xrm/2011/dxtools/")]
    [Microsoft.Xrm.Sdk.Client.RequestProxyAttribute("dxtools_SendCustomEmail")]
    [System.CodeDom.Compiler.GeneratedCodeAttribute("CrmSvcUtil", "6.0.0001.0061")]
    public partial class dxtools_SendCustomEmailRequest : Microsoft.Xrm.Sdk.OrganizationRequest
    {

        public string EmailTemplateName
        {
            get
            {
                if (this.Parameters.Contains("EmailTemplateName"))
                {
                    return ((string)(this.Parameters["EmailTemplateName"]));
                }
                else
                {
                    return default(string);
                }
            }
            set
            {
                this.Parameters["EmailTemplateName"] = value;
            }
        }

        public Microsoft.Xrm.Sdk.EntityReference FromSender
        {
            get
            {
                if (this.Parameters.Contains("FromSender"))
                {
                    return ((Microsoft.Xrm.Sdk.EntityReference)(this.Parameters["FromSender"]));
                }
                else
                {
                    return default(Microsoft.Xrm.Sdk.EntityReference);
                }
            }
            set
            {
                this.Parameters["FromSender"] = value;
            }
        }

        public Microsoft.Xrm.Sdk.EntityCollection ToRecipient
        {
            get
            {
                if (this.Parameters.Contains("ToRecipient"))
                {
                    return ((Microsoft.Xrm.Sdk.EntityCollection)(this.Parameters["ToRecipient"]));
                }
                else
                {
                    return default(Microsoft.Xrm.Sdk.EntityCollection);
                }
            }
            set
            {
                this.Parameters["ToRecipient"] = value;
            }
        }

        public Microsoft.Xrm.Sdk.EntityReference Regarding
        {
            get
            {
                if (this.Parameters.Contains("Regarding"))
                {
                    return ((Microsoft.Xrm.Sdk.EntityReference)(this.Parameters["Regarding"]));
                }
                else
                {
                    return default(Microsoft.Xrm.Sdk.EntityReference);
                }
            }
            set
            {
                this.Parameters["Regarding"] = value;
            }
        }

        public Microsoft.Xrm.Sdk.EntityReference RecordContext
        {
            get
            {
                if (this.Parameters.Contains("RecordContext"))
                {
                    return ((Microsoft.Xrm.Sdk.EntityReference)(this.Parameters["RecordContext"]));
                }
                else
                {
                    return default(Microsoft.Xrm.Sdk.EntityReference);
                }
            }
            set
            {
                this.Parameters["RecordContext"] = value;
            }
        }

        public string Attachments
        {
            get
            {
                if (this.Parameters.Contains("Attachments"))
                {
                    return ((string)(this.Parameters["Attachments"]));
                }
                else
                {
                    return default(string);
                }
            }
            set
            {
                this.Parameters["Attachments"] = value;
            }
        }

        public Boolean AllowDuplicateAttachments
        {
            get
            {
                if (this.Parameters.Contains("AllowDuplicateAttachments"))
                {
                    return ((Boolean)(this.Parameters["AllowDuplicateAttachments"]));
                }
                else
                {
                    return default(Boolean);
                }
            }
            set
            {
                this.Parameters["AllowDuplicateAttachments"] = value;
            }
        }

        public Boolean OnlyUseAttachmentsInTemplate
        {
            get
            {
                if (this.Parameters.Contains("OnlyUseAttachmentsInTemplate"))
                {
                    return ((Boolean)(this.Parameters["OnlyUseAttachmentsInTemplate"]));
                }
                else
                {
                    return default(Boolean);
                }
            }
            set
            {
                this.Parameters["OnlyUseAttachmentsInTemplate"] = value;
            }
        }


        public string StaticParameters
        {
            get
            {
                if (this.Parameters.Contains("StaticParameters"))
                {
                    return ((string)(this.Parameters["StaticParameters"]));
                }
                else
                {
                    return default(string);
                }
            }
            set
            {
                this.Parameters["StaticParameters"] = value;
            }
        }

        public dxtools_SendCustomEmailRequest()
        {
            this.RequestName = "dxtools_SendCustomEmail";
            this.EmailTemplateName = default(string);
            this.FromSender = default(Microsoft.Xrm.Sdk.EntityReference);
            this.ToRecipient = default(Microsoft.Xrm.Sdk.EntityCollection);
        }
    }
}

