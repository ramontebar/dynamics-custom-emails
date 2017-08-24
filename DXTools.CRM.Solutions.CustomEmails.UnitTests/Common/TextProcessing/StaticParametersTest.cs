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
    public class StaticParametersTest
    {
        [TestMethod]
        public void TestTextProcessor_StaticParameters()
        {
            IOrganizationService stubIOrganizationService = new StubIOrganizationService();

            Dictionary<string,string> contextDicitonary = new Dictionary<string,string>();
            contextDicitonary.Add("fullname","John Smith");
            contextDicitonary.Add("lastname", "Lock");
            contextDicitonary.Add("LASTNAME", "Smith");

            ITextProcessor crmTextProcessor = new TextProcessor(stubIOrganizationService);

            string textToBeProcessed = "Hello dear {fullname}. My lastname is {LASTNAME}";
            string expected = "Hello dear John Smith. My lastname is Lock";
            string actualResult = crmTextProcessor.Process(textToBeProcessed, contextDicitonary);

            Assert.AreEqual(expected, actualResult);  
        }

    }
}
