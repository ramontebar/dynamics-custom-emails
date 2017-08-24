using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Fakes;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Fakes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests
{
    public static class StubOrganizationServiceFactory
    {
        /// <summary>
        /// Simple helper method to set up Organisation Service
        /// </summary>
        /// <param name="recordContext"></param>
        /// <param name="attributeName"></param>
        /// <param name="attributeType"></param>
        /// <returns></returns>
        public static IOrganizationService SetupIOrganisationService(Entity recordContext, string attributeName, AttributeTypeCode attributeType)
        {
            StubIOrganizationService stubIOrganisationService = new StubIOrganizationService();
            stubIOrganisationService.RetrieveStringGuidColumnSet = (entityLogicalName, recordId, columnSet)
            =>
            {
                return recordContext;
            };

            stubIOrganisationService.ExecuteOrganizationRequest = (request)
            =>
            {
                if (request.RequestName == "RetrieveEntity")
                {
                    RetrieveEntityResponse retrieveEntityResponse = new RetrieveEntityResponse();
                    EntityMetadata entityMetadata = new EntityMetadata();
                    retrieveEntityResponse["EntityMetadata"] = entityMetadata;
                    return retrieveEntityResponse;
                }
                return null;
            };

            Microsoft.Xrm.Sdk.Metadata.Fakes.ShimEntityMetadata.AllInstances.AttributesGet =
                (item) =>
                {
                    return new AttributeMetadata[]
                    {
                        new StubAttributeMetadata(attributeType)
                        {
                            LogicalName = attributeName
                        }
                    };
                };

            return stubIOrganisationService;
        }

        /// <summary>
        /// Generic helper function to set up the Organisation service for several records and attributes
        /// </summary>
        /// <param name="stubData"></param>
        /// <returns></returns>
        public static IOrganizationService SetupIOrganisationService(List<StubData> stubData)
        {
            StubIOrganizationService stubIOrganisationService = new StubIOrganizationService();
            stubIOrganisationService.RetrieveStringGuidColumnSet = (entityLogicalName, recordId, columnSet)
            =>
            {
                var record = from r in stubData
                             where r.Record.LogicalName.Equals(entityLogicalName, StringComparison.InvariantCultureIgnoreCase)
                                   && r.Record.Id.Equals(recordId)
                             select r.Record;

                if (record != null)
                {
                    return record.FirstOrDefault();
                }
                else
                    return null;
            };

            stubIOrganisationService.ExecuteOrganizationRequest = (request)
            =>
            {
                if (request.RequestName == "RetrieveEntity")
                {
                    RetrieveEntityResponse retrieveEntityResponse = new RetrieveEntityResponse();
                    EntityMetadata entityMetadata = new EntityMetadata();
                    retrieveEntityResponse["EntityMetadata"] = entityMetadata;
                    return retrieveEntityResponse;
                }
                return null;
            };

            Microsoft.Xrm.Sdk.Metadata.Fakes.ShimEntityMetadata.AllInstances.AttributesGet =
                (item) =>
                {
                    var attributesMetadata = from r in stubData
                                             where r.Record.LogicalName.Equals(item.LogicalName, StringComparison.InvariantCultureIgnoreCase)
                                             select r.AttributesMetadata;

                    if (attributesMetadata != null)
                    {
                        return attributesMetadata.FirstOrDefault();
                    }
                    else
                        return null;

                };

            return stubIOrganisationService;
        }
    }
}
