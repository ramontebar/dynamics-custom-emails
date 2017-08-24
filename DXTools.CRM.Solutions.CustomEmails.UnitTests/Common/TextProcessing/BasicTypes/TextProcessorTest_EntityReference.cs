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
    public class TextProcessorTest_EntityReference
    {
        [TestMethod]
        public void TestTextProcessor_EntityReference()
        {
                Entity contact = new Entity("contact");
                contact["fullname"] = "Test Contact";
                contact.Id = Guid.NewGuid();

                Entity recordContext = new Entity("dxtools_payment");
                recordContext.Id = Guid.NewGuid();
                recordContext["dxtools_contactid"] = contact.ToEntityReference();
                recordContext["dxtools_paymentreference"] = "182929";
                

                IOrganizationService stubOrganisationService = SetupIOrganisationService(recordContext, contact);

                ITextProcessor dynTextProcessor = new DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing.TextProcessor(stubOrganisationService);

                string inputText = "{dxtools_contactid.fullname} is the owner of the payment {dxtools_paymentreference}";
                string expected = contact["fullname"].ToString() + " is the owner of the payment " + recordContext["dxtools_paymentreference"].ToString();

                string resultOuput = dynTextProcessor.Process(inputText, recordContext.ToEntityReference());

                Assert.AreEqual(expected, resultOuput);
            
        }

        private IOrganizationService SetupIOrganisationService(Entity recordContext, Entity contact)
        {
            StubIOrganizationService stubIOrganisationService = new StubIOrganizationService();
            stubIOrganisationService.RetrieveStringGuidColumnSet = (entityLogicalName, recordId, columnSet)
            =>
            {
                if (entityLogicalName == "dxtools_payment")
                    return recordContext;
                else if (entityLogicalName == "contact")
                    return contact;
                else
                    return null;
            };

            return stubIOrganisationService;
        }
    }
}
