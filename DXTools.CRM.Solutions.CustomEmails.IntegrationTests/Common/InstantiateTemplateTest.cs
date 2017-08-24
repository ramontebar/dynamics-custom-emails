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
    public class InstantiateTemplateTest
    {
        protected string TemplateName { get { return "DXTools Quote Test"; } }

        protected OrganizationService CrmOrganisationService { get; set; }

        [TestInitialize]
        public void InitialiseTests()
        {
            CrmConnection crmConnection = new CrmConnection("Xrm");
            crmConnection.ProxyTypesEnabled = true;
            CrmOrganisationService = new OrganizationService(crmConnection); 
        }

        [TestMethod]
        public void RetrieveExistingTemplate()
        {
            Entity template = TemplatesHelper.RetrieveTemplateByName(TemplateName, this.CrmOrganisationService);
            Assert.IsTrue(template.Contains("presentationxml"));

            InstantiateTemplateRequest instantiateTemplateRequest = new InstantiateTemplateRequest()
            {
                ObjectId = Systemuser.GetCallingUserID(this.CrmOrganisationService)
                  ,
                ObjectType = "systemuser"
                  ,
                TemplateId = template.Id
            };

            InstantiateTemplateResponse instantiateTemplateResponse = this.CrmOrganisationService.Execute(instantiateTemplateRequest) as InstantiateTemplateResponse;
            Assert.IsNotNull(instantiateTemplateResponse.EntityCollection.Entities[0]);

            Entity email = instantiateTemplateResponse.EntityCollection.Entities[0];
            Assert.IsTrue(email.Contains("description"));
        }

       
    }
}
