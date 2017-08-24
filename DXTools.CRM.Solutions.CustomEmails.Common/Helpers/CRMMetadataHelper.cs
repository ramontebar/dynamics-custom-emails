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
    public class CRMMetadataHelper
    {
        public static string GetEntityNameFromETC(int etc, IOrganizationService service)
        {
            EntityQueryExpression metadataQuery = new EntityQueryExpression();
            metadataQuery.Properties = new MetadataPropertiesExpression();
            metadataQuery.Properties.PropertyNames.Add("LogicalName");
            metadataQuery.Criteria.Conditions.Add(new MetadataConditionExpression("ObjectTypeCode", MetadataConditionOperator.Equals, etc));

            var retrieveMetadataChangesRequest = new RetrieveMetadataChangesRequest
            {
                Query = metadataQuery,
                ClientVersionStamp = null,
                DeletedMetadataFilters = DeletedMetadataFilters.OptionSet
            };

            RetrieveMetadataChangesResponse response = service.Execute(retrieveMetadataChangesRequest) as RetrieveMetadataChangesResponse;
            if (response.EntityMetadata.Count > 0)
                return response.EntityMetadata[0].LogicalName;
            else
                throw new Exception(string.Format("Entity with Object Type Code '{0}' couldn't be found", etc));

        }
    }
}
