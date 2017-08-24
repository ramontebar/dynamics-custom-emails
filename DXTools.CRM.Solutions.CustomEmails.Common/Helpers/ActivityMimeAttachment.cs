using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.Common.Helpers
{
    public static class ActivityMimeAttachment
    {
        public static EntityCollection RetrieveTemplatesActivityMimeAttachmentsByName(string[] attachments, IOrganizationService organisationService)
        {
            if (attachments == null || attachments.Length == 0)
                return null;

            QueryExpression query = new QueryExpression("activitymimeattachment");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet(new string[] { "filename" });
            query.Criteria.FilterOperator = LogicalOperator.And;
            query.Criteria.AddCondition("objecttypecode", ConditionOperator.Equal, "template");


            FilterExpression childFilter = query.Criteria.AddFilter(LogicalOperator.Or);
            for (int x = 0; x < attachments.Length; x++)
            {
                childFilter.AddCondition(
                    new ConditionExpression("filename", ConditionOperator.Equal, attachments[x])
                );
            }

            EntityCollection activitymimeattachmentCollection = organisationService.RetrieveMultiple(query);

            return activitymimeattachmentCollection;
        }
    }
}
