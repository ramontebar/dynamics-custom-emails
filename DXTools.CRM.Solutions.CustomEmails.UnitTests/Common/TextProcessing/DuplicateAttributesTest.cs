using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DXTools.CRM.Solutions.CustomEmails.Common;
using Microsoft.Xrm.Sdk.Fakes;
using Microsoft.Xrm.Sdk;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Fakes;
using Microsoft.QualityTools.Testing.Fakes;
using DXTools.CRM.Solutions.CustomEmails.Common.TextProcessing;

namespace DXTools.CRM.Solutions.CustomEmails.UnitTests.Common.TextProcessing
{
    [TestClass]
    public class DuplicateAttributesTest
    {
        [TestMethod]
        public void TestTextProcessor_DuplicateAttributes()
        {
            string fullname = "John Smith";

            Entity recordContext = new Entity("dxtools_payment");
            recordContext["fullname"] = fullname;
            recordContext.Id = Guid.NewGuid();

            IOrganizationService stubIOrganizationService = SetupIOrganisationService(recordContext);

            ITextProcessor crmTextProcessor = new TextProcessor(stubIOrganizationService);

            string textToBeProcessed = "Hello dear {fullname}. Now this is a duplicate {fullname}";
            string expected = string.Format("Hello dear {0}. Now this is a duplicate {0}", fullname);
            string actualResult = crmTextProcessor.Process(textToBeProcessed, recordContext.ToEntityReference());

            Assert.AreEqual(expected, actualResult);  
        }

        private IOrganizationService SetupIOrganisationService(Entity recordContext)
        {
            StubIOrganizationService stubIOrganisationService = new StubIOrganizationService();
            stubIOrganisationService.RetrieveStringGuidColumnSet = (entityLogicalName, recordId, columnSet)
            =>
            {
                    return recordContext;
            };

            return stubIOrganisationService;
        }

    }
}
