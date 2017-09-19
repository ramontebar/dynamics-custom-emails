using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Fakes;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Fakes;
using Microsoft.QualityTools.Testing.Fakes;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.Common.TextProcessing
{
    [TestClass]
    public class TextProcessorTest_Status
    {
        [TestMethod]
        public void TestTextProcessor_Status()
        {
            using (ShimsContext.Create())
            {
                Entity recordContext = new Entity("dxtools_payment");
                recordContext["statecode"] = new OptionSet;
                recordContext.Id = Guid.NewGuid();

                IOrganizationService stubOrganisationService = SetupIOrganisationService(recordContext, "createdon", AttributeTypeCode.DateTime);

                ITextProcessor dynTextProcessor = new TextProcessor(stubOrganisationService);

                string inputText = "Payment Status (statecode) is {statecode}.";
                string expected = "Your payment was created on " + DateTime.Now + ", thanks. ";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
            
        }

        private IOrganizationService SetupIOrganisationService(Entity recordContext, string attributeName, AttributeTypeCode attributeType)
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
    }
}
