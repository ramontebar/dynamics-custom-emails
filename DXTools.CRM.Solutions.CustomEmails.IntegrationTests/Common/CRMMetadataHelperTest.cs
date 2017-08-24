using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DXTools.CRM.Solutions.CustomEmails.Common;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using System.Collections.Generic;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using DXTools.CRM.Solutions.CustomEmails.Common.Helpers;

namespace DXTools.CRM.Solutions.CustomEmails.IntegrationTests.Common
{
    [TestClass]
    public class CRMMetadataHelperTest
    {
        protected OrganizationService CrmOrganisationService { get; set; }

        [TestInitialize]
        public void InitialiseTests()
        {
            CrmConnection crmConnection = new CrmConnection("Xrm");
            crmConnection.ProxyTypesEnabled = true;
            CrmOrganisationService = new OrganizationService(crmConnection); 
        }

        [TestMethod]
        public void GetEntityNameFromETC()
        {
            string expectedValue = "contact";
            string actualValue = CRMMetadataHelper.GetEntityNameFromETC(2, CrmOrganisationService);

            Assert.AreEqual(expectedValue, actualValue);
        }
    }
}
