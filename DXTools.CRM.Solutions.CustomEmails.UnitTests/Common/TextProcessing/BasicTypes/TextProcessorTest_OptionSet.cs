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
    public class TextProcessorTest_OptionSet
    {
        [TestMethod]
        public void TestTextProcessor_OptionSet()
        {
            using (ShimsContext.Create())
            {
                Entity recordContext = new Entity("dxtools_payment");
                recordContext.Id = Guid.NewGuid();
                recordContext["dxtools_paymentdirection"] = new OptionSetValue(503530000);

                IOrganizationService stubOrganisationService = SetupIOrganisationService(recordContext, 503530000);

                ISyntacticParser syntacticParser = new SyntacticParser(new LexicalParser());
                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "Payment direction is {dxtools_PaymentDirection}. Other option to get the same is {dxtools_paymentdirection.Label}, and the value is {dxtools_paymentdirection.Value}";
                string expected = "Payment direction is IN. Other option to get the same is IN, and the value is 503530000";

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            }
            
        }

        private IOrganizationService SetupIOrganisationService(Entity recordContext, int optionSetValue)
        {
            StubIOrganizationService stubIOrganisationService = new StubIOrganizationService();
            stubIOrganisationService.RetrieveStringGuidColumnSet = (entityLogicalName, recordId, columnSet)
            =>
            {
                if (entityLogicalName == "dxtools_payment")
                    return recordContext;
                else
                    return null;
            };

            stubIOrganisationService.ExecuteOrganizationRequest = (request)
            =>
            {
                if (request.RequestName == "RetrieveAttribute")
                {
                    RetrieveAttributeResponse response = new RetrieveAttributeResponse();
                    return response;
                }
                else
                    return null;
                    
            };

            Microsoft.Xrm.Sdk.Messages.Fakes.ShimRetrieveAttributeResponse.AllInstances.AttributeMetadataGet = (i)
                =>
            {
                OptionMetadataCollection optionMetadataCollection = new OptionMetadataCollection();
                OptionMetadata optionMetadata = new OptionMetadata(new Label(), optionSetValue);
                optionMetadata.Label.UserLocalizedLabel = new LocalizedLabel("IN",1033);
                optionMetadataCollection.Add(optionMetadata);

                PicklistAttributeMetadata picklistAttributeMetadata = new PicklistAttributeMetadata();
                picklistAttributeMetadata.OptionSet = new OptionSetMetadata(optionMetadataCollection);
                return picklistAttributeMetadata;
            };

            return stubIOrganisationService;
        }
    }
}
